"""Поиск таблиц с ИНН на листе (любое расположение: одна или несколько таблиц рядом)."""
import re
from collections import Counter

from utils import cell_val, norm_text, to_number
from compare_models import SideList

KEY_RE = re.compile(r"^\d{8,14}$")
TEXT_KEY_RE = re.compile(r"[\s,;(\[]*(?:инн|рма|рмя|inn)\D{0,5}?(\d{8,14})(?!\d)[)\]]?", re.I)


def key_of(v) -> str | None:
    """ИНН/РМА: 8–14 цифр (ведущие нули могли потеряться, если Excel хранит число)."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s if KEY_RE.match(s) else None


def key_in_text(v) -> str | None:
    """ИНН внутри текста: «ООО А (ИНН 123456789)» → 123456789."""
    m = TEXT_KEY_RE.search(norm_text(v)) if isinstance(v, str) else None
    return m.group(1) if m else None


def side_key(s, v) -> str | None:
    """ИНН из ячейки столбца ИНН таблицы (для «ИНН в тексте» — из текста)."""
    return key_in_text(v) if s.in_text else key_of(v)


def name_at(data, r: int, s) -> str:
    """Название в строке r таблицы s (без «(ИНН …)», если ИНН записан в тексте)."""
    v = str(cell_val(data, r, s.name_col) or "").strip() if s.name_col else ""
    return TEXT_KEY_RE.sub("", v).strip() if s.in_text else v


def match_key(k: str) -> str:
    """Ключ сопоставления без ведущих нулей."""
    return k.lstrip("0")


def _is_numbering(vals: list) -> bool:
    """Проверка, является ли столбец порядковым номером (1, 2, 3…)."""
    ints = [int(v) for v in vals if isinstance(v, (int, float)) and float(v).is_integer()]
    if not ints or len(ints) < 0.9 * len(vals):
        return False
    check_len = min(5, len(ints))
    return ints[:check_len] == list(range(ints[0], ints[0] + check_len))


def find_tables(data: list[list]):
    """Все таблицы с ИНН на листе: столбцы, где почти все значения — ИНН/РМА. → (шапка, [SideList])"""
    width = max((len(r) for r in data), default=0)
    key_cols, in_text = [], set()
    for c in range(1, width + 1):
        vals = [r[c - 1] for r in data if c - 1 < len(r) and r[c - 1] not in (None, "")]
        for find in (key_of, key_in_text):
            keys = [v for v in vals if find(v)]
            if len(keys) >= 10 and len(keys) >= 0.6 * len(vals):
                key_cols.append(c)
                if find is key_in_text:
                    in_text.add(c)
                break
    if not key_cols:
        return None

    first = min(r for r, row in enumerate(data, 1) for c in key_cols
                if c - 1 < len(row) and (key_in_text if c in in_text else key_of)(row[c - 1]))
    header = max(1, first - 1)
    for r in range(first - 1, 0, -1):
        row = data[r - 1]
        if any(isinstance(v, str) and len(v.strip()) > 1 and not norm_text(v).startswith(("хамаги", "итого", "всего"))
               for v in row):
            header = r
            break

    def used(c):
        return any(c - 1 < len(row) and row[c - 1] not in (None, "") for row in data[header - 1:])

    blocks, cur = [], []
    for c in range(1, width + 1):
        if used(c):
            cur.append(c)
        elif cur:
            blocks.append(cur)
            cur = []
    if cur:
        blocks.append(cur)

    def last_row(cols):
        return max((r for r, row in enumerate(data, 1) for c in cols
                    if c - 1 < len(row) and row[c - 1] not in (None, "")), default=0)

    sides = []
    for k in key_cols:
        block = next(b for b in blocks if k in b)
        sides.append(SideList(k, None, list(block), [], block=list(block), in_text=k in in_text))

    for b in blocks:
        if any(k in b for k in key_cols):
            continue
        lr = last_row(b)
        target = min(sides, key=lambda s: abs(last_row([s.key_col]) - lr))
        target.cols.extend(b)

    sides.sort(key=lambda s: min(s.block))
    for s in sides:
        s.header = header
        s.last_row = last_row(s.block)
        s.cols.sort()
        texts = Counter()
        for c in s.cols:
            if c in key_cols:           # столбцы ИНН — не суммы, даже если это числа
                continue
            col_vals = [row[c - 1] for row in data[header:] if c - 1 < len(row) and row[c - 1] not in (None, "")]
            if not col_vals:
                continue
            nums = [v for v in col_vals if to_number(v) is not None]
            strs = [v for v in col_vals if isinstance(v, str) and len(v.strip()) > 3 and not key_of(v)]
            if len(strs) > 0.5 * len(col_vals):
                texts[c] = len(strs)
            elif len(nums) > 0.5 * len(col_vals) and not _is_numbering(col_vals):
                s.value_cols.append(c)
        s.name_col = s.key_col if s.in_text else (texts.most_common(1)[0][0] if texts else None)
        for r in range(header + 1, len(data) + 1):
            k = side_key(s, data[r - 1][s.key_col - 1] if s.key_col - 1 < len(data[r - 1]) else None)
            if not k:
                continue
            s.all_rows.setdefault(match_key(k), []).append(r)
            if match_key(k) not in s.rows:
                s.rows[match_key(k)] = r
                s.keys[match_key(k)] = k
        years = Counter(y for c in s.value_cols for y in re.findall(r"20\d\d", str(cell_val(data, header, c) or "")))
        s.year = max(years) if years else None
    return header, sides
