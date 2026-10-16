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

    def test_several_sheets_each_own_columns(self):
        """Два листа одного файла: реестр за январь и список — каждый со своими столбцами."""
        path = os.path.join(tempfile.mkdtemp(), "два листа.xlsx")
        wb = Workbook()
        wb.active.title = "Январь"
        wb.active.append(["ИНН плательщика", "ИНН получателя", "Сумма"])
        wb.active.append([INNS[0], INNS[1], 50])
        ws = wb.create_sheet("Список")
        ws.append(["Название", "РМА"])
        ws.append(["Орг", INNS[0]])
        wb.save(path)
        maps = [{"sheet": "Январь", "head": 0, "row0": 0, "col0": 0, "kind": REGISTRY,
                 "cols": {"payer_inn": 0, "rec_inn": 1, "amount": [2]}},
                {"sheet": "Список", "head": 0, "row0": 0, "col0": 0, "kind": AS_LIST, "cols": {"inn": 1, "name": 0},
                 "sums": []}]
        reg, lists, _ = scan_files([path], {path: maps})["raw"]
        self.assertEqual((reg.payers[INNS[0]].total, [e.name for e in lists.entities(0, 1)]), (50.0, ["Орг"]))

    def test_table_not_from_a1(self):
        """Таблица с C3: строка заголовков и номера строк — как в Excel."""
        path = os.path.join(tempfile.mkdtemp(), "c3.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Л"
        for r, row in enumerate([["ИНН плательщика", "ИНН получателя", "Сумма"], [INNS[0], INNS[1], 5]], 3):
            for c, v in enumerate(row, 3):
                ws.cell(r, c, v)
        wb.save(path)
        p = propose(path)
        self.assertEqual((p["guess"]["head"], p["sheets"]["Л"]["row0"], p["sheets"]["Л"]["col0"]), (2, 0, 2))
        from registry_read import read_registry
        self.assertEqual(next(read_registry(path)[0][0][2]).row, 4)

    def test_dialog_typed_columns(self):
        """Окно: пусто — «Готово» не нажимается; колонки вписываются от руки (как в Olimv3)."""
        from PySide6.QtWidgets import QApplication
        from ui_file_dialog import FileDialog
        app = QApplication.instance() or QApplication([])  # noqa: F841
        path = _save([["ИНН", "Название", "План", "Факт"]] + [[inn, f"Орг {i}", 1, 2] for i, inn in enumerate(INNS)])
        d = FileDialog(None, propose(path), 1, 1)
        ok = d.buttons.button(d.buttons.StandardButton.Ok)
        self.assertFalse(ok.isEnabled())                            # ничего не подставлено
        d.edits["inn"].setText("a2-a5")
        d.edits["sums"].setText("c, d")
        self.assertTrue(ok.isEnabled())
        m = d.mapping()[0]
        self.assertEqual((m["kind"], m["head"], m["last"], m["cols"]["inn"], m["sums"]), (AS_LIST, 0, 5, 0, [2, 3]))
        d.edits["inn"].setText("a5-b9")
        self.assertFalse(ok.isEnabled())
        self.assertIn("одной колонке", d.warn.text())

if __name__ == "__main__":
    unittest.main()
