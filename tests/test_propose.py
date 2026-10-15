import os
import sys
import tempfile
import unittest
from datetime import date

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook  # noqa: E402

from analysis_list_map import AS_LIST  # noqa: E402
from analysis_propose import REGISTRY, propose  # noqa: E402
from analysis_scan import scan_files  # noqa: E402

INNS = [f"{100000000 + i * 7919:09d}" for i in range(12)]


def _save(rows, title="Л"):
    path = os.path.join(tempfile.mkdtemp(), "f.xlsx")
    wb = Workbook()
    wb.active.title = title
    for r in rows:
        wb.active.append(r)
    wb.save(path)
    return path


class ProposeTest(unittest.TestCase):
    def test_registry_guess(self):
        path = _save([["Отчёт"], ["№ док", "Дата док", "Плательщик", "ИНН плательщика", "Получатель", "ИНН получателя",
                                  "Сумма"], [1, date(2024, 1, 5), "А", INNS[0], "Б", INNS[1], 10.0]])
        g = propose(path)["guess"]
        self.assertEqual((g["kind"], g["head"], g["cols"]["payer_inn"], g["sums"]), (REGISTRY, 1, 3, [6]))

    def test_list_guess_and_manual_sums(self):
        """Список: догадка — ИНН и название; пользователь выбрал суммировать Дебет и Кредит."""
        path = _save([["ИНН", "Контрагент", "Дебет", "Кредит"]] + [[inn, f"Орг {i}", 100, 1] for i, inn in enumerate(INNS)])
        p = propose(path)
        self.assertEqual((p["guess"]["kind"], p["guess"]["cols"]), (AS_LIST, {"inn": 0, "name": 1}))
        mapping = {"sheet": "Л", "head": 0, "row0": 0, "col0": 0, "cols": {"inn": 0, "name": 1}, "kind": AS_LIST,
                   "sums": [2, 3]}
        _, lists, _ = scan_files([path], {path: mapping})["raw"]
        self.assertEqual({e.sums[0] for e in lists.entities(0, 1)}, {101.0})

    def test_registry_several_sum_columns(self):
        path = _save([["ИНН плательщика", "ИНН получателя", "Основная", "Пеня"], [INNS[0], INNS[1], 100, 5]])
        mapping = {"sheet": "Л", "head": 0, "row0": 0, "col0": 0, "kind": REGISTRY,
                   "cols": {"payer_inn": 0, "rec_inn": 1, "amount": [2, 3]}}
        reg, _, _ = scan_files([path], {path: mapping})["raw"]
        self.assertEqual(reg.payers[INNS[0]].total, 105.0)


if __name__ == "__main__":
    unittest.main()
