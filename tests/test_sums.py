"""Суммы: не числа в столбцах сумм — на отдельный лист; отбор платежей реестра по основанию / получателю."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook  # noqa: E402

from analysis_scan import scan_files  # noqa: E402
from utils import not_a_number  # noqa: E402

INNS = [f"{100000000 + i * 7919:09d}" for i in range(4)]


def _save(rows):
    path = os.path.join(tempfile.mkdtemp(), "f.xlsx")
    wb = Workbook()
    wb.active.title = "Л"
    for r in rows:
        wb.active.append(r)
    wb.save(path)
    return path


class SumsTest(unittest.TestCase):
    def test_not_a_number(self):
        self.assertEqual([not_a_number(v) for v in ("нет", "1.234.56", "-", "—", "", None, "1 234,56", 5)],
                         [True, True, False, False, False, False, False, False])

    def test_list_bad_amount_sheet(self):
        path = _save([["ИНН", "Название", "Сумма"], [INNS[0], "А", "1.234.56"], [INNS[1], "Б", 10]])
        maps = {path: [{"sheet": "Л", "head": 0, "row0": 0, "col0": 0, "kind": "list", "cols": {"inn": 0, "name": 1},
                        "sums": [2]}]}
        lists, _ = scan_files([path], maps)["raw"]
        self.assertEqual(lists.details["bad_num_list"], [["f.xlsx", "«Л»", 2, "C", "1.234.56", "А"]])



if __name__ == "__main__":
    unittest.main()
