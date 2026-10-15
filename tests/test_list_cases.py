"""Непохожие на наши файлы: заголовки ниже 30-й строки, перевёрнутая таблица, несколько организаций в ячейке,
список без ИНН (по названию)."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook  # noqa: E402

from analysis_list_map import AS_LIST  # noqa: E402
from analysis_model import BY_NAME, SHARED  # noqa: E402
from analysis_propose import REGISTRY, propose  # noqa: E402
from analysis_reconcile import reconcile  # noqa: E402
from analysis_scan import build, scan_files  # noqa: E402

INNS = [f"{100000000 + i * 7919:09d}" for i in range(4)]


def _save(name, rows):
    path = os.path.join(tempfile.mkdtemp(), name)
    wb = Workbook()
    wb.active.title = "Л"
    for r in rows:
        wb.active.append(r)
    wb.save(path)
    return path


def _list_map(head, cols, sums=(), **extra):
    return [dict({"sheet": "Л", "head": head, "row0": 0, "col0": 0, "kind": AS_LIST, "cols": cols,
                  "sums": list(sums)}, **extra)]


def _orgs(path, maps):
    _, lists, _ = scan_files([path], {path: maps})["raw"]
    return {e.name: e for e in lists.entities(0, len(lists.parts))}


class ListCasesTest(unittest.TestCase):
    def test_header_below_row_30(self):
        path = _save("низко.xlsx", [["пояснение"]] * 119 + [["ИНН", "Название"], [INNS[0], "Орг"]])
        self.assertGreaterEqual(len(propose(path)["sheets"]["Л"]["rows"]), 121)
        self.assertEqual(list(_orgs(path, _list_map(119, {"inn": 0, "name": 1}))), ["Орг"])

    def test_transposed_table(self):
        """Организации по столбцам: строка 1 — ИНН, строка 2 — названия, строка 3 — суммы."""
        path = _save("перевёрнута.xlsx", [["ИНН"] + INNS[:2], ["Название", "А", "Б"], ["Сумма", 10, 20]])
        maps = [{"sheet": "Л", "head": 0, "row0": 0, "col0": 0, "kind": AS_LIST, "transpose": True,
                 "cols": {"inn": 0, "name": 1}, "sums": [2]}]
        orgs = _orgs(path, maps)
        self.assertEqual({n: (e.inn, e.sums[0]) for n, e in orgs.items()}, {"А": (INNS[0], 10.0), "Б": (INNS[1], 20.0)})

    def test_several_orgs_in_one_cell(self):
        """«ИНН1, ИНН2» и «А, Б» в одной строке — каждой организации вся сумма строки и пометка «общая строка»."""
        path = _save("общая.xlsx", [["ИНН", "Название", "Сумма"], [f"{INNS[0]}, {INNS[1]}", "А, Б", 100],
                                    [INNS[2], "В", 5]])
        orgs = _orgs(path, _list_map(0, {"inn": 0, "name": 1}, [2]))
        self.assertEqual({n: e.sums[0] for n, e in orgs.items()}, {"А": 100.0, "Б": 100.0, "В": 5.0})
        self.assertEqual((SHARED in orgs["А"].flags, SHARED in orgs["В"].flags), (True, False))

    def test_no_inn_match_by_name(self):
        """Список без ИНН — организация находится в реестре по названию (без кавычек и формы ООО)."""
        lst = _save("без ИНН.xlsx", [["Название", "План"], ['ООО "Альфа"', 50], ["Бета", 7]])
        reg = _save("реестр.xlsx", [["Плательщик", "ИНН плательщика", "Получатель", "ИНН получателя", "Сумма"],
                                    ["Альфа ООО", INNS[0], "Бюджет", INNS[3], 30]])
        maps = {lst: _list_map(0, {"inn": None, "name": 0}, [1], by_name=True),
                reg: [{"sheet": "Л", "head": 0, "row0": 0, "col0": 0, "kind": REGISTRY,
                       "cols": {"payer": 0, "payer_inn": 1, "rec": 2, "rec_inn": 3, "amount": [4]}}]}
        data = build(*scan_files([lst, reg], maps)["raw"])
        _, heads, rows = reconcile(data)
        alpha = next(r for r in rows if "Альфа" in r.name)
        self.assertEqual((alpha.values, alpha.paid, alpha.to[0][0].lower()), ([50.0], [30.0], "бюджет"))
        self.assertTrue(all(BY_NAME in e.flags for e in data.entities if e.role == "организация"))
        self.assertIsNone(data.compare)         # реестр + список — разные вещи: «Что случилось» не считается


if __name__ == "__main__":
    unittest.main()
