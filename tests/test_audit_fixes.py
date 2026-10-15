"""Найдено проверкой на файлах аудитора: шапка реестра ниже 20-й строки, проверка ИНН в списках при ручных
столбцах, минус в скобках, План и Факт — отдельно, а не суммой."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook  # noqa: E402

from analysis_list_map import AS_LIST  # noqa: E402
from analysis_propose import REGISTRY  # noqa: E402
from analysis_scan import build, scan_files  # noqa: E402
from utils import to_number  # noqa: E402

INNS = [f"{100000000 + i * 7919:09d}" for i in range(3)]


def _save(rows):
    path = os.path.join(tempfile.mkdtemp(), "f.xlsx")
    wb = Workbook()
    wb.active.title = "Л"
    for r in rows:
        wb.active.append(r)
    wb.save(path)
    return path


def _list(head, cols, sums, **extra):
    return [dict({"sheet": "Л", "head": head, "row0": 0, "col0": 0, "kind": AS_LIST, "cols": cols, "sums": sums}, **extra)]


class AuditFixesTest(unittest.TestCase):
    def test_registry_header_below_row_20(self):
        path = _save([["Отчёт"]] + [[None]] * 48 + [["ИНН плательщика", "ИНН получателя", "Сумма"], [INNS[0], INNS[1], 7]])
        maps = [{"sheet": "Л", "head": 49, "row0": 0, "col0": 0, "kind": REGISTRY,
                 "cols": {"payer_inn": 0, "rec_inn": 1, "amount": [2]}}]
        reg, _, _ = scan_files([path], {path: maps})["raw"]
        self.assertEqual(reg.payers[INNS[0]].total, 7.0)

    def test_list_inn_errors_with_manual_columns(self):
        path = _save([["ИНН", "Название", "Сумма"], [INNS[0], "А", 1], [INNS[0], "А", 1], [12345678, "Потерян ноль", 1],
                      ["0", "Ноль", 1], [None, "Без ИНН", 1], [None, "Итого", 5]])
        _, lists, _ = scan_files([path], {path: _list(0, {"inn": 0, "name": 1}, [2])})["raw"]
        bad = {r[4]: (r[5], r[6]) for r in lists.details["bad_list"]}
        self.assertEqual(bad["Потерян ноль"], ("8 цифр вместо 9 — возможно, потерян 0 в начале (должно быть 012345678)",
                                               "012345678"))
        self.assertEqual((bad["Ноль"][0], bad["Без ИНН"][0], "Итого" in bad), ("вместо ИНН стоит 0 (не указан)",
                                                                               "ИНН не указан", False))
        self.assertEqual([r[2] for r in lists.details["dup_row"]], ["2, 3"])

    def test_minus_in_brackets(self):
        self.assertEqual((to_number("(1 000,00)"), to_number("(abc)")), (-1000.0, None))

    def test_plan_and_fact_separately(self):
        """«Не складывать»: План и Факт — две части (столбцы) вместо одной суммы."""
        path = _save([["ИНН", "Учреждение", "План", "Факт"], [INNS[0], "А", 1000, 900], [INNS[1], "Б", 500, 600]])
        maps = {path: _list(0, {"inn": 0, "name": 1}, [2, 3], separate=True)}
        data = build(*scan_files([path], maps)["raw"])
        self.assertEqual([p.label for p in data.parts], ["«Л»: План", "«Л»: Факт"])
        self.assertEqual({e.name: e.sums for e in data.entities}, {"А": [1000.0, 900.0], "Б": [500.0, 600.0]})


if __name__ == "__main__":
    unittest.main()
