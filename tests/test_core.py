import unittest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from parser import parse_purpose, parse_purpose_ex, to_number
from inn_keys import key_of, match_key
from excel_io import parse_cell, cell_name, gib_code


class TestParser(unittest.TestCase):
    def test_single_gib(self):
        res, longs = parse_purpose_ex("Оплата по счету ГИБ 2211-140")
        self.assertEqual(res, {"2211": 140.0})
        self.assertEqual(longs, [])

    def test_formatted_amount(self):
        res, longs = parse_purpose_ex("ГИБ 2211 - 1 500,50 сомони")
        self.assertEqual(res, {"2211": 1500.50})

    def test_long_code(self):
        res, longs = parse_purpose_ex("ГИБ 22110-500.00")
        self.assertEqual(res, {"2211": 500.0})
        self.assertEqual(longs, ["22110"])

    def test_sub_items_sum_into_one_column(self):
        # как в образце: «ГИБ 22111-600,00» и «ГИБ 22112-550,00» — 1150,00 в одном столбце 2211
        res, longs = parse_purpose_ex("ГИБ 22111-600,00 ГИБ 22112-550,00")
        self.assertEqual(res, {"2211": 1150.0})
        self.assertEqual(longs, ["22111", "22112"])

    def test_three_digit_code(self):
        # как в образце: код из 3 цифр — тоже ГИБ
        res, longs = parse_purpose_ex("ГИБ 221-10")
        self.assertEqual(res, {"221": 10.0})
        self.assertEqual(longs, [])

    def test_inn_not_taken_as_amount(self):
        self.assertEqual(parse_purpose("ГИБ 2211-140 020012345"), {"2211": 140.0})

    def test_gib_header_codes(self):
        self.assertEqual(gib_code(2211), "2211")
        self.assertEqual(gib_code("221 Хизмат"), "221")
        self.assertIsNone(gib_code("111110 Хазина"))      # счёт листа «мухо» — не ГИБ

    def test_multiple_gib(self):
        res, _ = parse_purpose_ex("ГИБ 2111-100 и ГИБ 2211-200.50")
        self.assertEqual(res, {"2111": 100.0, "2211": 200.50})

    def test_summed_same_gib(self):
        res = parse_purpose("ГИБ 2111-100 ГИБ 2111-50")
        self.assertEqual(res, {"2111": 150.0})

    def test_no_gib(self):
        self.assertEqual(parse_purpose("Оплата услуг без кода"), {})

    def test_to_number(self):
        self.assertEqual(to_number("1 250,50"), 1250.50)
        self.assertEqual(to_number(100), 100.0)
        self.assertIsNone(to_number(""))
        self.assertIsNone(to_number(None))


class TestCompareHelpers(unittest.TestCase):
    def test_key_of(self):
        self.assertEqual(key_of("012345678"), "012345678")
        self.assertEqual(key_of(12345678.0), "12345678")
        self.assertIsNone(key_of("123"))
        self.assertIsNone(key_of("abc"))

    def test_match_key(self):
        self.assertEqual(match_key("0012345678"), "12345678")


class TestExcelHelpers(unittest.TestCase):
    def test_cell_name(self):
        self.assertEqual(cell_name(5, 2), "B5")
        self.assertEqual(parse_cell("B5"), (5, 2))

    def test_gib_code(self):
        self.assertEqual(gib_code("2111 Музди маош"), "2111")
        self.assertEqual(gib_code(2211), "2211")
        self.assertIsNone(gib_code("Итого"))


if __name__ == "__main__":
    unittest.main()
