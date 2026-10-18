"""Глубокий анализ: одно окно «Листы и ИНН» на все файлы; пустые файлы пропускаются; 3 файла по несколько листов —
в своде все листы, сумм нет, лист «без ИНН» скопирован как есть."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook, load_workbook  # noqa: E402
from PySide6.QtWidgets import QApplication, QDialog  # noqa: E402

import ui_analysis_flow as flow  # noqa: E402
from analysis_jobs import make_report  # noqa: E402
from analysis_propose import propose_files  # noqa: E402
from analysis_scan import scan_files  # noqa: E402
from sheet_choice import choices, to_mappings  # noqa: E402

INNS = [f"{100000000 + i * 7919:09d}" for i in range(12)]


def _book(name: str, sheets: dict) -> str:
    path = os.path.join(tempfile.mkdtemp(), name)
    wb = Workbook()
    for k, (title, rows) in enumerate(sheets.items()):
        ws = wb.active if k == 0 else wb.create_sheet()
        ws.title = title
        for r in rows:
            ws.append(r)
    wb.save(path)
    return path


def _orgs(sum_: int, inns=INNS) -> list:
    return [["Название", "ИНН", "Сумма"]] + [[f"Орг {i}", inn, sum_] for i, inn in enumerate(inns)]


class _Dialog:
    """Окно «Листы и ИНН» без экрана: принимает то, что нашла программа (или «Отмена»)."""
    accept = True

    def __init__(self, win, props):
        self.props = props

    def exec(self):
        return QDialog.Accepted if _Dialog.accept else QDialog.Rejected

    def mappings(self):
        return to_mappings(self.props, choices(self.props))[0]


class AnalysisFlowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def setUp(self):
        self.saved = (flow.SheetsDialog, flow.QMessageBox.warning, flow.QMessageBox.information)
        flow.SheetsDialog = _Dialog
        flow.QMessageBox.warning = flow.QMessageBox.information = lambda *a: self.notes.append(a[2])
        self.notes, self.used = [], None
        _Dialog.accept = True

    def tearDown(self):
        flow.SheetsDialog, flow.QMessageBox.warning, flow.QMessageBox.information = self.saved

    def _ask(self, props):
        win = type("Win", (), {})()
        flow._ask_sheets(win, props, lambda used: setattr(self, "used", used))
        return win

    def test_empty_file_skipped(self):
        a, b = _book("a.xlsx", {"Л": _orgs(1)}), _book("b.xlsx", {"Л": []})
        win = self._ask(propose_files([a, b]))
        self.assertEqual(self.used, [a])
        self.assertEqual(list(win.an_mappings), [a])
        self.assertIn("b.xlsx", self.notes[0])

    def test_all_files_empty(self):
        self._ask(propose_files([_book("a.xlsx", {"Л": []})]))
        self.assertIsNone(self.used)
        self.assertIn("пустые", self.notes[0])

    def test_cancel(self):
        _Dialog.accept = False
        self._ask(propose_files([_book("a.xlsx", {"Л": _orgs(1)})]))
        self.assertIsNone(self.used)

    def test_three_files_many_sheets(self):
        """3 файла по 2 листа + лист без ИНН: каждый лист — своя колонка, сумм нет, «Файлов» по файлам."""
        paths = [_book(f"{m}.xlsx", {"Банк": _orgs(5), "Касса": _orgs(7, INNS[:10])}) for m in ("Янв", "Фев")]
        paths.append(_book("Мар.xlsx", {"Банк": _orgs(9), "Прочее": [["заметка", "текст"], ["ещё", 1]]}))
        props = propose_files(paths)
        rows = choices(props)
        self.assertEqual([r["found"] for r in rows], [True, True, True, True, True, False])
        rows[-1]["raw"] = True                                  # «Прочее» — без ИНН, как есть
        maps, problems = to_mappings(props, rows)
        self.assertEqual(problems, [])
        raw = scan_files(paths, maps)["raw"]
        out = os.path.join(tempfile.mkdtemp(), "r.xlsx")
        res = make_report(out, raw, maps)
        wb = load_workbook(out)
        head = [c.value for c in wb["Анализ"][5]]
        self.assertEqual(len([h for h in head if "Банк" in str(h) or "Касса" in str(h)]), 5)   # 5 листов с ИНН
        for gone in ("Всего", "Что случилось", "Изменение"):
            self.assertNotIn(gone, head)
        org = {r[1]: r for r in wb["Анализ"].iter_rows(min_row=6, values_only=True)}
        self.assertEqual(org[INNS[0]][head.index("Файлов")], 3)
        self.assertEqual(org[INNS[11]][head.index("Есть в файлах")], "Янв.xlsx, Фев.xlsx, Мар.xlsx")
        self.assertEqual(wb.sheetnames[:2], ["Одна орг. во всех файлах", "Анализ"])   # главный свод — первым
        self.assertEqual([r[:3] for r in res["raw"]], [("Мар.xlsx", "Прочее", "Мар · Прочее")])
        self.assertEqual([list(r) for r in wb["Мар · Прочее"].iter_rows(values_only=True)], [["заметка", "текст"], ["ещё", 1]])


if __name__ == "__main__":
    unittest.main()
