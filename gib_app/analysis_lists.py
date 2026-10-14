"""Обычные файлы со списками организаций (таблицы с ИНН на любых листах) → части и организации
глубокого анализа. Ошибки — те же правила, что в проверке ИНН (inn_check)."""
import os

from analysis_model import BAD_INN, DUP_ROW, ORG, Entity, Part
from compare_finder import find_tables
from inn_check import _name_key, check_inn
from utils import cell_val, to_num_safe


class ListCollector:
    """Собирает таблицы с ИНН из нескольких файлов: одна организация (ИНН) — одна строка."""

    def __init__(self):
        self.parts = []
        self.orgs = {}          # ключ ИНН (без нулей впереди) → {"inn", "names", "sums", "counts", "dup"}
        self.details = []       # строки «Ошибки подробно»
        self.bad = set()        # ИНН (как в файле) с ошибкой

    def add_file(self, path: str, sheets: dict) -> int:
        """→ сколько таблиц с ИНН нашлось в файле."""
        fname, before = os.path.basename(path), len(self.parts)
        for sheet, data in sheets.items():
            found = find_tables(data)
            if found:
                for s in found[1]:
                    self._add_side(fname, sheet, data, s, len(found[1]))
        if len(self.parts) > before:
            self._add_errors(fname, sheets)
        return len(self.parts) - before

    def _add_side(self, fname, sheet, data, s, on_sheet):
        i = len(self.parts)
        where = f"«{sheet}» {s.span}" if on_sheet > 1 else f"«{sheet}»"
        metric = s.value_cols[0] if s.value_cols else None          # первый столбец сумм таблицы
        self.parts.append(Part(f"{s.year} · {where}" if s.year else where, fname, "list",
                               (int(s.year) if s.year else 9999, 0, i), has_sums=metric is not None))
        for k, rows in s.all_rows.items():
            o = self.orgs.setdefault(k, {"inn": s.keys[k], "names": [], "sums": {}, "counts": {}, "dup": False})
            o["names"] += [n for n in (str(cell_val(data, r, s.name_col) or "").strip() if s.name_col else ""
                                       for r in rows) if n]
            o["sums"][i] = round(sum(to_num_safe(cell_val(data, r, metric)) for r in rows), 2) if metric else 0.0
            o["counts"][i] = len(rows)
            o["dup"] = o["dup"] or len(rows) > 1

    def _add_errors(self, fname, sheets):
        res = check_inn(sheets)
        if not res:
            return
        self.details += [[DUP_ROW, fname, f"{t.label}, строки {', '.join(map(str, rows))}", inn, name,
                          None, None, None, f"записан {len(rows)} раза"] for t, inn, name, rows in res.duplicates]
        self.details += [[BAD_INN, fname, f"{t.label}, строка {row}", raw or "(пусто)", name, None, None, None,
                          problem + (f"; должно быть {fix}" if fix else "")]
                         for t, row, raw, name, problem, fix in res.invalid]
        self.bad |= {raw for _, row, raw, *_ in res.invalid if raw}

    def entities(self, offset: int, n_parts: int) -> list:
        """Организации; их суммы — в позициях offset… общей таблицы из n_parts частей."""
        out = []
        for o in self.orgs.values():
            sums, counts = [0.0] * n_parts, [0] * n_parts
            for i, v in o["sums"].items():
                sums[offset + i], counts[offset + i] = v, o["counts"][i]
            e = Entity(o["inn"], o["names"][0] if o["names"] else "", ORG, sums, counts,
                       names=max(1, len({_name_key(n) for n in o["names"]})))
            if o["inn"] in self.bad:
                e.flags.add(BAD_INN)
            if o["dup"]:
                e.flags.add(DUP_ROW)
            out.append(e)
        return out
