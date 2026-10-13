"""Проверка таблиц с ИНН / РМА: дубликаты, один ИНН — разные названия, неправильные ИНН.

Проверяются все таблицы с ИНН на всех листах файла. Файл не меняется.
"""
import re
from collections import Counter
from dataclasses import dataclass, field

from compare_finder import find_tables, match_key
from constants import SUM_ROW_LABELS as TOTAL_LABELS
from utils import cell_val as _cell, norm_text


@dataclass
class Table:
    sheet: str
    span: str
    side: str
    count: int

    @property
    def label(self):
        part = f" — {self.side}" if self.side else ""
        return f"«{self.sheet}»{part} (столбцы {self.span})"


@dataclass
class CheckResult:
    tables: list = field(default_factory=list)       # [Table]
    duplicates: list = field(default_factory=list)   # (таблица, ИНН, название, [строки])
    names: list = field(default_factory=list)        # (ИНН, [(название, таблица, строка)])
    invalid: list = field(default_factory=list)      # (таблица, строка, что стоит, название, проблема, как должно быть)

    @property
    def problems(self):
        return len(self.duplicates) + len(self.names) + len(self.invalid)


def _name_key(name):
    """Название для сравнения: без кавычек, знаков и лишних пробелов, таджикские буквы → русские."""
    return re.sub(r"[^0-9a-zа-яё]+", "", norm_text(name))


def _raw_key(v):
    """ИНН как написан в ячейке: строка без пробелов по краям; число без «.0»."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def check_inn(sheets):
    """Проверить все таблицы с ИНН в книге → CheckResult (или None, если ИНН нет нигде)."""
    res = CheckResult()
    by_inn = {}      # ИНН -> [(название, таблица, строка)]
    shown = {}       # ИНН -> как он записан в файле (с нулями в начале)
    for sheet, data in sheets.items():
        found = find_tables(data)
        if not found:
            continue
        header, sides = found
        for n, s in enumerate(sides):
            side = ("левая часть" if n == 0 else "правая часть") if len(sides) == 2 else (
                f"таблица {n + 1}" if len(sides) > 2 else "")
            t = Table(sheet, s.span, side, len(s.rows))
            res.tables.append(t)

            def name_at(r):
                return str(_cell(data, r, s.name_col) or "").strip() if s.name_col else ""

            # дубликаты внутри таблицы
            for k, rows in s.all_rows.items():
                if len(rows) > 1:
                    res.duplicates.append((t, s.keys[k], name_at(rows[0]), rows))
                shown.setdefault(k, s.keys[k])
                for r in rows:
                    by_inn.setdefault(k, []).append((name_at(r), t, r))

            # неправильные ИНН: длина — по самой частой длине в этом столбце
            lengths = Counter(len(s.keys[k]) for k in s.keys)
            normal = lengths.most_common(1)[0][0] if lengths else 9
            for r in range(header + 1, s.last_row + 1):
                row = data[r - 1] if r <= len(data) else []
                block_vals = [row[c - 1] for c in s.block if c - 1 < len(row)]
                if any(norm_text(v).startswith(TOTAL_LABELS) for v in block_vals if isinstance(v, str)):
                    continue
                raw_v = _cell(data, r, s.key_col)
                raw = _raw_key(raw_v)
                name = name_at(r)
                if raw is None:
                    if len(name) > 3:
                        res.invalid.append((t, r, "", name, "ИНН не указан", ""))
                    continue
                digits = re.sub(r"\s+", "", raw)
                if not digits.isdigit():
                    res.invalid.append((t, r, raw, name, "в ИНН есть буквы или знаки", ""))
                elif digits != raw:
                    res.invalid.append((t, r, raw, name, "в ИНН есть пробелы", digits))
                elif len(digits) < normal and isinstance(raw_v, (int, float)) and len(digits.zfill(normal)) == normal:
                    res.invalid.append((t, r, raw, name, f"потерялся ноль в начале ({len(digits)} цифр вместо {normal})",
                                        digits.zfill(normal)))
                elif len(digits) != normal:
                    res.invalid.append((t, r, raw, name, f"{len(digits)} цифр вместо {normal}", ""))

    if not res.tables:
        return None
    # один ИНН — разные названия (в одной таблице или на разных листах)
    for k, entries in by_inn.items():
        variants = {}
        for name, t, r in entries:
            if name:
                variants.setdefault(_name_key(name), (name, t, r))
        if len(variants) > 1:
            res.names.append((shown.get(k, k), list(variants.values())))
    res.names.sort(key=lambda x: -len(x[1]))
    return res
