import os
import sys
import unittest
from datetime import date

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from registry_analysis import gone_and_new, jumps, party_rows, search  # noqa: E402
import tempfile  # noqa: E402

from openpyxl import Workbook  # noqa: E402

from registry_columns import guess_header, registry_columns  # noqa: E402
from registry_read import Payment, inn_text, read_registry, to_payment  # noqa: E402
from registry_stats import Registry, period_label  # noqa: E402
from app_paths import ANALYSIS_FOLDER  # noqa: E402
from report_paths import registry_report_path  # noqa: E402
from report_registry_errors import inn_in_name  # noqa: E402

HEAD = ["№ Документа", "Дата документа", "Дата учета", "Плательщик", "ИНН Плат", "Счет Плательщика", "Сумма",
        "БИК плат", "Получатель", "ИНН получ", "Счет Получателя", "БИК Получ", "КоррСчет получ", "ТипДокумента",
        "Основание"]


def pay(row, payer_inn, rec_inn, amount, d=date(2024, 1, 10), doc=None):
    return Payment(row, doc or str(row), d, "Орг " + payer_inn, payer_inn, "Получ " + rec_inn, rec_inn, amount)


class TestRegistryRead(unittest.TestCase):
    def test_columns_by_header(self):
        cols = registry_columns(HEAD)
        self.assertEqual((cols["payer"], cols["payer_inn"], cols["amount"], cols["rec"], cols["rec_inn"]),
                         (3, 4, 6, 8, 9))
        self.assertIsNone(registry_columns(["№", "Основание", "Расходы"]))

    def test_inn_text_restores_zero(self):
        self.assertEqual(inn_text(10002120), "010002120")
        self.assertEqual(inn_text(" 040000011 "), "040000011")

    def test_totals_row_skipped(self):
        cols = registry_columns(HEAD)
        self.assertIsNone(to_payment(5, [None] * 6 + [3219817.0] + [None] * 8, cols))
        p = to_payment(2, ["1", date(2024, 1, 4), None, "А", "010002120", None, 200, None, "Б", "251010603"], cols)
        self.assertEqual((p.row, p.amount, p.rec_inn), (2, 200.0, "251010603"))


class TestOtherDocuments(unittest.TestCase):
    def test_tajik_and_other_headers(self):
        cols = registry_columns(["№", "Сана", "Номи пардохткунанда", "РМА пардохткунанда", "Маблағ",
                                 "Гиранда", "РМА гиранда"])
        self.assertEqual((cols["payer_inn"], cols["amount"], cols["rec_inn"], cols["date"]), (3, 4, 6, 1))
        cols = registry_columns(["Наименование плательщика", "ИНН плательщика", "Сумма платежа"])
        self.assertEqual((cols["payer"], cols["payer_inn"], cols["amount"], cols["rec_inn"]), (0, 1, 2, None))

    def test_one_side_document(self):
        cols = registry_columns(["Получатель", "ИНН получ", "Сумма"])
        p = to_payment(2, ["Орг", "040000011", 100], cols)
        self.assertEqual((p.payer_inn, p.rec_inn), (None, "040000011"))
        r = Registry()
        f = r.start_file("x.xlsx", "Лист1")
        r.add(f, p)
        self.assertEqual((list(r.recs), r.payers, r.bad_inn), (["040000011"], {}, []))

    def test_guess_header(self):
        g = guess_header([["Отчёт"], [None], ["Колонка1", "Сумма", "Колонка3"], [1, 100, "x"]])
        self.assertEqual((g["head"], g["cols"]["amount"], g["heads"][0]), (2, 1, "Колонка1"))

    def _book(self, sheets):
        path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
        wb = Workbook()
        wb.remove(wb.active)
        for name, rows in sheets:
            ws = wb.create_sheet(name)
            for row in rows:
                ws.append(row)
        wb.save(path)
        return path

    def test_all_registry_sheets_are_read(self):
        head = ["Плательщик", "ИНН Плат", "Сумма", "Получатель", "ИНН получ"]
        path = self._book([("Январь", [head, ["А", "111111111", 10, "Б", "222222222"]]),
                           ("Свод", [["Итого", 10]]),
                           ("Февраль", [head, ["А", "111111111", 20, "Б", "222222222"]])])
        sheets, guess = read_registry(path)
        self.assertEqual(([s[0] for s in sheets], guess), (["Январь", "Февраль"], None))
        self.assertEqual([p.amount for _, _, rows in sheets for p in rows], [10.0, 20.0])

    def test_user_columns(self):
        path = self._book([("Лист1", [["К1", "К2", "К3"], ["111111111", "Орг", 55]])])
        sheets, guess = read_registry(path)
        self.assertEqual((sheets, guess["sheet"], guess["head"]), ([], "Лист1", 0))
        mapping = {"sheet": "Лист1", "head": 0, "cols": {"payer_inn": 0, "payer": 1, "amount": 2}}
        sheets, _ = read_registry(path, mapping)
        p = next(sheets[0][2])
        self.assertEqual((p.payer_inn, p.payer, p.amount, p.rec_inn, p.row), ("111111111", "Орг", 55.0, None, 2))


class TestRegistryStats(unittest.TestCase):
    def setUp(self):
        r = self.reg = Registry()
        f1 = r.start_file("a.xlsx", "Лист1")
        for i in range(200):
            r.add(f1, pay(i + 2, "111111111", "999999999", 10.0, date(2024, 1 + i % 4, 5)))
        r.add(f1, pay(300, "222222222", "999999999", 100.0))
        r.add(f1, pay(301, "333333333", "999999999", 50.0, d=date(2004, 3, 27)))     # странная дата
        r.add(f1, pay(302, "0", "999999999", 5.0))                                 # плохой ИНН
        r.add(f1, pay(303, "222222222", "999999999", 100.0, doc="300"))            # двойной платёж
        f2 = r.start_file("b.xlsx", "Лист1")
        for i in range(200):
            r.add(f2, pay(i + 2, "111111111", "999999999", 10.0, date(2024, 5 + i % 4, 5)))
        r.add(f2, pay(300, "333333333", "999999999", 500.0, date(2024, 6, 1)))
        r.add(f2, pay(301, "444444444", "999999999", 7.0, date(2024, 6, 1)))
        r.finish()

    def test_periods(self):
        self.assertEqual([self.reg.files[f].label for f in self.reg.order], ["Янв–Апр 2024", "Май–Авг 2024"])
        self.assertEqual(period_label([(2024, 1), (2025, 2)]), "Янв 2024 – Фев 2025")

    def test_errors(self):
        self.assertEqual([p.row for _, p in self.reg.odd_dates], [301])
        self.assertEqual([(p.row, side) for _, p, side, _ in self.reg.bad_inn], [(302, "плательщик")])
        self.assertEqual([(p.row, first) for _, p, first in self.reg.dups], [(303, (0, 300))])

    def test_gone_new_jumps(self):
        rows = party_rows(self.reg.payers, self.reg)
        gone, new = gone_and_new(rows)
        self.assertEqual([r.inn for r in gone], ["222222222"])
        self.assertEqual([r.inn for r in new], ["444444444"])
        jj = jumps(rows, self.reg)
        self.assertEqual([(j[0].inn, j[3], j[4]) for j in jj], [("333333333", 50.0, 500.0)])

    def test_search_restores_zero(self):
        self.reg.payers["011111111"] = self.reg.payers.pop("111111111")
        res = search(self.reg, "11111111")
        self.assertEqual(res["payer"].sums, [2000.0, 2000.0])
        self.assertIsNone(res["rec"])

    def test_inn_in_name(self):
        self.assertEqual(inn_in_name("Сарраесати хазинадори ИНН 010010058"), "010010058")
        self.assertEqual(inn_in_name("счёт 20204972221022300242"), "")


class TestPaths(unittest.TestCase):
    def test_report_goes_to_analysis_folder(self):
        path = registry_report_path("Янв–Дек 2024")
        self.assertEqual(os.path.basename(os.path.dirname(path)), ANALYSIS_FOLDER)
        self.assertTrue(os.path.basename(path).startswith("Свод по ИНН (Янв–Дек 2024)"))


if __name__ == "__main__":
    unittest.main()
