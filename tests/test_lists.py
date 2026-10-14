import os
import sys
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from analysis_lists import ListCollector, payments_like  # noqa: E402

INNS = [f"{100000000 + i * 7919:09d}" for i in range(12)]


def _sums(head, rows):
    lists = ListCollector()
    lists.add_file("t.xlsx", {"Л": [head] + rows})
    return {e.inn: sum(e.sums) for e in lists.entities(0, len(lists.parts))}


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

    def test_two_inn_columns_unknown_headers(self):
        """Незнакомые заголовки, два столбца ИНН: ИНН не считается суммой, файл похож на реестр."""
        head = ["Кто платит", "Код плательщика", "Кому", "Код получателя", "Деньги"]
        rows = [[f"А {i}", inn, f"Б {i}", INNS[-1 - i], 10] for i, inn in enumerate(INNS)]
        self.assertTrue(payments_like({"Л": [head] + rows}))
        self.assertFalse(payments_like({"Л": [["ИНН", "Название"]] + [[i, "А"] for i in INNS]}))
        self.assertEqual(set(_sums(head, rows).values()), {20.0})    # 10 как плательщик + 10 как получатель


if __name__ == "__main__":
    unittest.main()
