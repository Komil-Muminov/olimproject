import os
import sys
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from report_errors import error_counts  # noqa: E402


class ErrorCountsTest(unittest.TestCase):
    def test_same_numbers_as_sheets_by_file(self):
        tail = ["А", "111111111", "Б", "222222222", ""]
        details = {
            "dup": [[1, "M1.xlsx", 5, "1", None, 10.0, "платёж"] + tail, [1, "M2.xlsx", 9, "1", None, 10.0, "повтор"] + tail],
            "bad_pay": [["M1.xlsx", 3], ["M1.xlsx", 4], ["M2.xlsx", 7]],
            "odd": [],
            "same_name": [[1, "А"], [1, "А"], [2, "Б"], [2, "Б"]],
        }
        self.assertEqual(error_counts(details), [
            ("Двойной платёж", {"M2.xlsx": 1}, 1),                 # пара — в файле, где повтор
            ("Неправильный ИНН", {"M1.xlsx": 2, "M2.xlsx": 1}, 3),
            ("Одно название — разные ИНН", {}, 2),                 # группы, без разбивки по файлам
        ])


if __name__ == "__main__":
    unittest.main()
