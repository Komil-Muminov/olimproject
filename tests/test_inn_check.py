import unittest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from inn_check import check_inn, _name_key, _raw_key


class TestInnCheck(unittest.TestCase):
    def test_name_key(self):
        self.assertEqual(_name_key("«Ширкат»"), _name_key("  Ширкат  "))
        self.assertEqual(_name_key("«Корхона №1»"), _name_key("Корхона 1"))

    def test_raw_key(self):
        self.assertEqual(_raw_key("  012345678  "), "012345678")
        self.assertEqual(_raw_key(12345678.0), "12345678")
        self.assertIsNone(_raw_key(""))
        self.assertIsNone(_raw_key(None))

    def test_check_inn_empty_sheet(self):
        sheets = {"Лист1": [["A", "B"], ["C", "D"]]}
        self.assertIsNone(check_inn(sheets))


if __name__ == "__main__":
    unittest.main()
