"""Глубокий анализ, несколько файлов: «Отмена» на одном файле и пустой файл не останавливают всё;
сравнения «было → стало» нет — на листе «Анализ» видно, в каких файлах есть организация."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook, load_workbook  # noqa: E402
from PySide6.QtWidgets import QApplication, QDialog  # noqa: E402

import ui_analysis_flow as flow  # noqa: E402
from analysis_list_map import AS_LIST  # noqa: E402
from analysis_jobs import make_report  # noqa: E402
from analysis_scan import auto_mode, scan_files  # noqa: E402

INNS = [f"{100000000 + i * 7919:09d}" for i in range(3)]
MAP = [{"sheet": "Л", "head": 0, "row0": 0, "col0": 0, "kind": AS_LIST, "cols": {"inn": 0, "name": 1}, "sums": [2]}]


def _save(name, amount):
    path = os.path.join(tempfile.mkdtemp(), name)
    wb = Workbook()
    wb.active.title = "Л"
    for r in (["ИНН", "Название", "Сумма"], [INNS[0], "А", amount], [INNS[1], "Б", 10]):
        wb.active.append(r)
    wb.save(path)
    return path


def _raw(n):
    paths = [_save(f"{i + 1}.xlsx", 100 * (i + 1)) for i in range(n)]
    return scan_files(paths, {p: MAP for p in paths})["raw"]


class _Dialog:
    """Окно «Как читать файл» без экрана: «ОК» — для путей из ok, иначе «Отмена»."""
    ok = set()

    def __init__(self, win, prop, k, n):
        self.path = prop["path"]

    def exec(self):
        return QDialog.Accepted if self.path in _Dialog.ok else QDialog.Rejected

    def mapping(self):
        return MAP


class AnalysisFlowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def setUp(self):
        self.saved = (flow.FileDialog, flow._skip, flow.QMessageBox.warning, flow.QMessageBox.information)
        flow.FileDialog = _Dialog
        flow.QMessageBox.warning = flow.QMessageBox.information = lambda *a: self.notes.append(a[2])
        self.notes, self.used = [], None

    def tearDown(self):
        flow.FileDialog, flow._skip, flow.QMessageBox.warning, flow.QMessageBox.information = self.saved

    def _ask(self, props, skip):
        flow._skip = lambda win, name, rest: skip
        win = type("Win", (), {})()
        flow._ask_files(win, props, lambda used: setattr(self, "used", used))
        return win

    def test_cancel_one_file_keeps_others(self):
        props = [{"path": p, "sheets": [1]} for p in ("a.xlsx", "b.xlsx", "c.xlsx")]
        _Dialog.ok = {"a.xlsx", "c.xlsx"}
        win = self._ask(props, skip=True)
        self.assertEqual(self.used, ["a.xlsx", "c.xlsx"])
        self.assertEqual(set(win.an_mappings), {"a.xlsx", "c.xlsx"})
        self.assertIn("b.xlsx — пропущен", self.notes[0])

    def test_cancel_and_stop(self):
        _Dialog.ok = {"a.xlsx"}
        self._ask([{"path": p, "sheets": [1]} for p in ("a.xlsx", "b.xlsx")], skip=False)
        self.assertIsNone(self.used)

    def test_empty_file_skipped(self):
        _Dialog.ok = {"a.xlsx", "c.xlsx"}
        props = [{"path": "a.xlsx", "sheets": [1]}, {"path": "b.xlsx", "sheets": []}, {"path": "c.xlsx", "sheets": [1]}]
        self._ask(props, skip=False)
        self.assertEqual(self.used, ["a.xlsx", "c.xlsx"])
        self.assertIn("b.xlsx — пустой", self.notes[0])

    def test_all_files_empty(self):
        self._ask([{"path": "a.xlsx", "sheets": []}], skip=False)
        self.assertIsNone(self.used)
        self.assertIn("Нечего анализировать", self.notes[0])

    def test_three_files_no_comparison(self):
        """Три файла: на «Анализ» нет «Что случилось» и «Изменение», есть «Есть в файлах» и «Файлов»."""
        raw = _raw(3)
        out = os.path.join(tempfile.mkdtemp(), "r.xlsx")
        make_report(out, raw, auto_mode(raw))
        ws = load_workbook(out)["Анализ"]
        head = [c.value for c in ws[5]]
        for gone in ("Что случилось", "Изменение", "Изменение, %"):
            self.assertNotIn(gone, head)
        rows = {r[1]: r for r in ws.iter_rows(min_row=6, values_only=True)}
        files = head.index("Файлов")
        self.assertEqual((rows[INNS[0]][files], rows[INNS[0]][files - 1]), (3, "1.xlsx, 2.xlsx, 3.xlsx"))


if __name__ == "__main__":
    unittest.main()
