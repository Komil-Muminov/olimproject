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

    def test_sheets_dialog(self):
        """Окно «Листы и ИНН»: ИНН нашёлся сам — ✓; не нашёлся — «Анализировать» не нажимается, пока не вписать
        колонку или не отметить «Без ИНН» (лист — в свод как есть)."""
        from PySide6.QtWidgets import QApplication
        from analysis_list_map import RAW
        from ui_sheets_dialog import FOUND, NOT_FOUND, SheetsDialog
        app = QApplication.instance() or QApplication([])  # noqa: F841
        path = os.path.join(tempfile.mkdtemp(), "f.xlsx")
        wb = Workbook()
        wb.active.title = "Банк"
        for r in [["Название", "ИНН"]] + [[f"Орг {i}", inn] for i, inn in enumerate(INNS)]:
            wb.active.append(r)
        wb.create_sheet("Прочее").append(["просто", "текст"])
        wb.save(path)
        d = SheetsDialog(None, [propose(path)])
        ok = d.buttons.button(d.buttons.StandardButton.Ok)
        bank, other = d.widgets
        self.assertEqual((bank["inn"].text(), bank["name"].text(), bank["status"].text()), ("B", "A", FOUND))
        self.assertEqual(other["status"].text(), NOT_FOUND)
        self.assertFalse(ok.isEnabled())
        other["inn"].setText("a")                                   # в колонке A нет ИНН — понятная ошибка
        self.assertIn("нет ИНН", other["status"].text())
        other["raw"].setChecked(True)
        self.assertTrue(ok.isEnabled())
        m = d.mappings()[path]
        self.assertEqual([(x["sheet"], x["kind"]) for x in m], [("Банк", AS_LIST), ("Прочее", RAW)])
        self.assertEqual((m[0]["cols"]["inn"], m[0]["cols"]["name"], m[0]["sums"]), (1, 0, []))
        bank["use"].setChecked(False)                                # лист не берём
        other["use"].setChecked(False)
        self.assertFalse(ok.isEnabled())


if __name__ == "__main__":
    unittest.main()
