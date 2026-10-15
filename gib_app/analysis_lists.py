"""Обычные файлы со списками организаций (таблицы с ИНН на любых листах) → части и организации
глубокого анализа. Ошибки — те же правила, что в проверке ИНН (inn_check)."""
import os

from analysis_list_map import FIND, mapped_table, side_guess
from analysis_model import BAD_INN, DUP_ROW, ORG, Entity, Part
from compare_finder import find_tables, name_at
from inn_check import _name_key, check_inn
from utils import cell_val, norm_text, to_num_safe

# деньги в двух столбцах (ушло / пришло) — считаются оба
TWO_SIDED = ("дебет", "кредит", "debit", "credit", "приход", "расход")


def _named(data, s, words) -> list:
    """Столбцы с числами, чей заголовок начинается с одного из words."""
    return [c for c in s.value_cols if norm_text(cell_val(data, s.header, c) or "").startswith(words)]


def money_cols(data, s) -> list:
    """Столбцы сумм таблицы: Дебет и Кредит (Приход/Расход) — оба; иначе столбец «Сумма», иначе первый с числами."""
    if s.fixed_money:
        return s.value_cols
    both = _named(data, s, TWO_SIDED)
    return both if len(both) > 1 else (_named(data, s, FIND["amount"])[:1] or s.value_cols[:1])


def money_unclear(sheets: dict) -> dict | None:
    """Несколько столбцов с числами и ни один не «Сумма» (и не Дебет/Кредит) — какой считать деньгами,
    скажет пользователь. → догадка для окна выбора столбцов или None."""
    for sheet, data in sheets.items():
        found = find_tables(data)
        for s in (found[1] if found else []):
            if len(s.value_cols) > 1 and len(_named(data, s, TWO_SIDED)) < 2 and not _named(data, s, FIND["amount"]):
                return side_guess(data, sheet, s)
    return None


def payments_like(sheets: dict) -> bool:
    """Два столбца ИНН в одной таблице (плательщик и получатель) — похоже на реестр платежей."""
    for data in sheets.values():
        found = find_tables(data)
        if found and len({tuple(s.block) for s in found[1]}) < len(found[1]):
            return True
    return False


class ListCollector:
    """Собирает таблицы с ИНН из нескольких файлов: одна организация (ИНН) — одна строка."""

    def __init__(self):
        self.parts = []
        self.orgs = {}          # ключ ИНН (без нулей впереди) → {"inn", "names", "sums", "counts", "dup"}
        self.details = {"dup_row": [], "bad_list": []}      # группы листов ошибок в таблицах
        self.bad = set()        # ИНН (как в файле) с ошибкой

    def add_file(self, path: str, sheets: dict, mapping: dict = None) -> int:
        """→ сколько таблиц с ИНН нашлось в файле. mapping — столбцы, указанные пользователем."""
        fname, before = os.path.basename(path), len(self.parts)
        if mapping:                     # столбцы указал пользователь — читается только этот лист
            s = mapped_table(sheets[mapping["sheet"]], mapping)
            if s.all_rows:
                self._add_side(fname, mapping["sheet"], sheets[mapping["sheet"]], s, 1)
            return len(self.parts) - before
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
        metric = money_cols(data, s)
        self.parts.append(Part(f"{s.year} · {where}" if s.year else where, fname, "list",
                               (int(s.year) if s.year else 9999, 0, i), has_sums=bool(metric)))
        for k, rows in s.all_rows.items():
            o = self.orgs.setdefault(k, {"inn": s.keys[k], "names": [], "sums": {}, "counts": {}, "dup": False})
            o["names"] += [n for n in (name_at(data, r, s) for r in rows) if n]
            o["sums"][i] = round(sum(to_num_safe(cell_val(data, r, c)) for r in rows for c in metric), 2)
            o["counts"][i] = len(rows)
            o["dup"] = o["dup"] or len(rows) > 1

    def _add_errors(self, fname, sheets):
        res = check_inn(sheets)
        if not res:
            return
        self.details["dup_row"] += [[fname, t.label, ", ".join(map(str, rows)), inn, name, f"{len(rows)} раза"]
                                    for t, inn, name, rows in res.duplicates]
        self.details["bad_list"] += [[fname, t.label, row, raw or "(пусто)", name, problem, fix]
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
