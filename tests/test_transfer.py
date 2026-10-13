import os
import sys
import unittest
from datetime import datetime

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from statement import doc_text, statement_header, to_date  # noqa: E402
from target_layout import _amount, payment_key  # noqa: E402


class TestStatement(unittest.TestCase):
    def test_header_found(self):
        data = [["", "Выписка"], ["", "№ док-та", "Дата Док-а", "Дата Учета", "Основание", "Доходы", "Расходы"]]
        self.assertEqual(statement_header(data), (2, 5, 7, 2, 3))

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


if __name__ == "__main__":
    unittest.main()
