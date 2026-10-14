import os
import sys
import tempfile
import unittest

from openpyxl import Workbook

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from analysis_list_map import AS_LIST  # noqa: E402
from analysis_lists import ListCollector, payments_like  # noqa: E402
from analysis_scan import scan_files  # noqa: E402

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

    def test_inn_inside_text(self):
        rows = [[f"ООО «Альфа {i}» (ИНН {inn})", 100] for i, inn in enumerate(INNS)]
        rows[1][0] = f"Бета, РМА: {INNS[1]}"
        lists = ListCollector()
        lists.add_file("t.xlsx", {"Л": [["Контрагент", "Сумма"]] + rows})
        ents = {e.inn: e for e in lists.entities(0, 1)}
        self.assertEqual(set(ents), set(INNS))
        self.assertEqual((ents[INNS[0]].name, ents[INNS[1]].name, ents[INNS[0]].sums), ("ООО «Альфа 0»", "Бета", [100.0]))
        self.assertEqual(lists.details["bad_list"], [])                   # текст вокруг ИНН — не ошибка


class AskInnTest(unittest.TestCase):
    """ИНН не нашёлся сам (строк мало) — спросить столбец; по указанному столбцу таблица читается."""

    def test_small_list(self):
        path = os.path.join(tempfile.mkdtemp(), "мало.xlsx")
        wb = Workbook()
        wb.active.append(["Наименование", "ИНН", "Сумма"])
        for i, inn in enumerate(INNS[:5]):
            wb.active.append([f"Орг {i}", inn, 10])
        wb.active.append(["Итого", None, 50])
        wb.save(path)
        res = scan_files([path])
        self.assertEqual((res["needs_inn"], res["guess"]["cols"]), (path, {"inn": 1, "name": 0, "amount": 2}))
        mapping = {"sheet": res["guess"]["sheet"], "head": 0, "cols": res["guess"]["cols"], "kind": AS_LIST}
        reg, lists, _ = scan_files([path], {path: mapping})["raw"]
        ents = lists.entities(0, 1)
        self.assertEqual((len(ents), sum(e.sums[0] for e in ents), ents[0].name), (5, 50.0, "Орг 0"))


if __name__ == "__main__":
    unittest.main()
