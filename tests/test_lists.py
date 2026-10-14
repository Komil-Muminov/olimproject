import os
import sys
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from analysis_lists import ListCollector  # noqa: E402

INNS = [f"{100000000 + i * 7919:09d}" for i in range(12)]


def _sums(head, rows):
    lists = ListCollector()
    lists.add_file("t.xlsx", {"Л": [head] + rows})
    return {e.inn: e.sums[0] for e in lists.entities(0, 1)}


class ListMoneyTest(unittest.TestCase):
    def test_debit_and_credit_both_counted(self):
        rows = [[inn, f"Организация {i}", 100 * (i + 1) if i % 2 else None, None if i % 2 else 10 * (i + 1)]
                for i, inn in enumerate(INNS)]
        sums = _sums(["ИНН", "Контрагент", "Дебет", "Кредит"], rows)
        self.assertEqual(sums[INNS[0]], 10.0)        # только кредит
        self.assertEqual(sums[INNS[1]], 200.0)       # только дебет

    def test_one_money_column(self):
        rows = [[inn, f"Организация {i}", 5, 7] for i, inn in enumerate(INNS)]
        sums = _sums(["ИНН", "Наименование", "Сумма", "Остаток"], rows)
        self.assertEqual(sums[INNS[0]], 5.0)         # только первый столбец сумм


if __name__ == "__main__":
    unittest.main()
