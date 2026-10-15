"""Суммы: не числа в столбцах сумм — на отдельный лист; отбор платежей реестра по основанию / получателю."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook  # noqa: E402

from analysis_scan import build, scan_files  # noqa: E402
from registry_read import Payment, payment_filter  # noqa: E402
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


def _reg(**extra):
    return [dict({"sheet": "Л", "head": 0, "row0": 0, "col0": 0, "kind": "registry",
                  "cols": {"payer_inn": 0, "rec": 1, "rec_inn": 2, "purpose": 3, "amount": [4]}}, **extra)]


ROWS = [["ИНН плательщика", "Получатель", "ИНН получателя", "Основание", "Сумма"],
        [INNS[0], "Бюджет", INNS[3], "Дивиденд 30% за 2023", 100], [INNS[0], "ООО Свет", INNS[1], "за свет", 50],
        [INNS[2], "Бюджет", INNS[3], "налог", "нет"], [INNS[2], "ООО Свет", INNS[1], "за свет", "—"]]


class SumsTest(unittest.TestCase):
    def test_not_a_number(self):
        self.assertEqual([not_a_number(v) for v in ("нет", "1.234.56", "-", "—", "", None, "1 234,56", 5)],
                         [True, True, False, False, False, False, False, False])

    def test_registry_bad_amount_sheet(self):
        path = _save(ROWS)
        data = build(*scan_files([path], {path: _reg()})["raw"])
        self.assertEqual([(r[1], r[5]) for r in data.details["bad_num_pay"]], [(4, "нет")])   # прочерк — не ошибка

    def test_list_bad_amount_sheet(self):
        path = _save([["ИНН", "Название", "Сумма"], [INNS[0], "А", "1.234.56"], [INNS[1], "Б", 10]])
        maps = {path: [{"sheet": "Л", "head": 0, "row0": 0, "col0": 0, "kind": "list", "cols": {"inn": 0, "name": 1},
                        "sums": [2]}]}
        _, lists, _ = scan_files([path], maps)["raw"]
        self.assertEqual(lists.details["bad_num_list"], [["f.xlsx", "«Л»", 2, "C", "1.234.56", "А"]])

    def test_filter_by_purpose_and_recipient(self):
        p = lambda rec, inn, purpose: Payment(1, "", None, "", INNS[0], rec, inn, 1.0, purpose)
        keep = payment_filter({"f_purpose": "дивиденд, 30%", "f_rec": INNS[3]})
        self.assertEqual([keep(p("Бюджет", INNS[3], "Дивиденд")), keep(p("Бюджет", INNS[3], "налог")),
                          keep(p("ООО", INNS[1], "дивиденд"))], [True, False, False])
        self.assertIsNone(payment_filter({"f_purpose": " , "}))
        path = _save(ROWS)
        reg, _, _ = scan_files([path], {path: _reg(f_purpose="дивиденд")})["raw"]
        self.assertEqual((reg.payers[INNS[0]].total, INNS[2] in reg.payers), (100.0, False))


if __name__ == "__main__":
    unittest.main()
