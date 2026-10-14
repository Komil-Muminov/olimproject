import os
import sys
import unittest
from datetime import datetime

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from gib_columns import target_cols  # noqa: E402
from statement import Payment, _payments, doc_text, statement_header, to_date  # noqa: E402
from target_layout import _amount, payment_key, split_new  # noqa: E402
from utils import find_col  # noqa: E402


class TestStatement(unittest.TestCase):
    def test_header_found(self):
        data = [["", "Выписка"], ["", "№ док-та", "Дата Док-а", "Дата Учета", "Основание", "Доходы", "Расходы"]]
        self.assertEqual(statement_header(data), (2, {"purpose": 4, "amount": 6, "doc": 1, "date": 2}))

    def test_unknown_headers_and_user_columns(self):
        data = [["Тавсиф", "Маблаг", "Ракам"], ["Пардохт ГИБ 2211-100,00", 100, "7"], ["Итого", 100, None]]
        self.assertIsNone(statement_header(data))
        pays = _payments(data, 1, {"purpose": 0, "amount": 1, "doc": 2, "date": None})
        self.assertEqual([(p.row, p.amount, p.doc, p.gibs) for p in pays], [(2, 100.0, "7", {"2211": 100.0})])

    def test_target_columns(self):
        cols = target_cols(["р/т", "Номгӯи корхонаҳо", "№ ҳуҷҷат", "Сана", "", "Ҳамагӣ", "Расходы"])
        self.assertEqual(cols, {"purpose": 1, "rash": 6, "doc": 2, "date": 3, "num": 0})

    def test_date_and_doc(self):
        self.assertEqual(to_date("22.12.2025"), datetime(2025, 12, 22))
        self.assertEqual(doc_text("0002309"), doc_text(2309.0))

    def test_amount_from_balance_formula(self):
        self.assertEqual(_amount("=ROUND(F6-18000,2)", 0), 18000.0)
        self.assertEqual(_amount("=ROUND(F6-(-5.5),2)", 0), -5.5)
        self.assertEqual(_amount(None, "3350"), 3350.0)

    def test_payment_key_ignores_case_and_spaces(self):
        self.assertEqual(payment_key("ГИБ 2211-100  оплата", 100.0, "0001"),
                         payment_key("гиб 2211-100 оплата", 100.0, 1.0))

    def test_split_new_counts_duplicates(self):
        a, b = Payment(1, "ГИБ 2211-10", 10.0), Payment(2, "ГИБ 2211-10", 10.0)
        c = Payment(3, "ГИБ 2214-5", 5.0)
        seen = {payment_key("гиб 2211-10", 10.0): 1}
        new, skipped = split_new([a, b, c], seen, use_doc=False)
        self.assertEqual(([p.row for p in new], skipped), ([2, 3], 1))

    def test_find_col(self):
        heads = ["р/т", "основание", "расходы", "расходы 2"]
        self.assertEqual(find_col(heads, ("расход",)), 3)
        self.assertEqual(find_col(heads, ("расход",), 3), 4)
        self.assertIsNone(find_col(heads, ("сана",)))


if __name__ == "__main__":
    unittest.main()
