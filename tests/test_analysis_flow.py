"""Глубокий анализ, несколько файлов: у каждого карточка; не взятый и пустой файл не мешают остальным;
больше двух частей — пользователь сам выбирает «Было» и «Стало»."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook  # noqa: E402
from PySide6.QtWidgets import QApplication, QDialogButtonBox  # noqa: E402

import ui_analysis_flow as flow  # noqa: E402
from analysis_list_map import AS_LIST  # noqa: E402
from analysis_scan import auto_mode, build, scan_files  # noqa: E402
from ui_pair_dialog import PairDialog, needs_pair  # noqa: E402

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


class AnalysisFlowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def _win(self, props):
        """Экран анализа без окна: карточки и перехват запуска фоновой задачи."""
        from ui_file_card import FileCard
        win = type("Win", (), {})()
        win.worker, win.an_run_btn = None, type("B", (), {"setEnabled": lambda s, v: setattr(s, "on", v)})()
        win.an_file_cards = [FileCard(p) for p in props]
        from PySide6.QtWidgets import QLineEdit
        win.an_min, win.an_max = QLineEdit(), QLineEdit()                  # порог суммы — пусто
        self.started = None
        flow._run = lambda w, func, *args, done: setattr(self, "started", args)
        return win

    def setUp(self):
        self.saved = flow._run

    def tearDown(self):
        flow._run = self.saved

    def test_only_included_files_go_to_analysis(self):
        a, b, c = (_save(n, 5) for n in ("a.xlsx", "b.xlsx", "c.xlsx"))
        from analysis_propose import propose
        win = self._win([propose(p) for p in (a, b, c)])
        for card in win.an_file_cards:
            card.edits["inn"].setText("a")
        win.an_file_cards[1].use.setChecked(False)              # b не берём
        flow.run(win)
        paths, mappings = self.started
        self.assertEqual((paths, set(mappings)), ([a, c], {a, c}))

    def test_not_ready_until_inn_given(self):
        from analysis_propose import propose
        win = self._win([propose(_save("a.xlsx", 5))])
        flow.run(win)
        self.assertIsNone(self.started)                          # ИНН не вписан — «Анализ» не запускается

    def test_empty_file_card(self):
        from ui_file_card import FileCard
        card = FileCard({"path": "пустой.xlsx", "sheets": {}})
        self.assertEqual((card.included(), card.use.isEnabled(), card.problem()), (False, False, ""))

    def test_pair_asked_only_for_more_than_two_parts(self):
        raw2, raw3 = _raw(2), _raw(3)
        self.assertFalse(needs_pair(raw2, auto_mode(raw2)))
        self.assertTrue(needs_pair(raw3, auto_mode(raw3)))

    def test_pair_dialog_choice(self):
        raw = _raw(3)
        dlg = PairDialog(None, raw, auto_mode(raw))
        self.assertEqual(dlg.pair(), (0, 2))                 # по умолчанию: первая → последняя
        dlg.now.setCurrentIndex(0)
        self.assertFalse(dlg.buttons.button(QDialogButtonBox.Ok).isEnabled())
        dlg.now.setCurrentIndex(1)
        self.assertTrue(dlg.buttons.button(QDialogButtonBox.Ok).isEnabled())
        data = build(*raw, auto_mode(raw), dlg.pair())
        self.assertEqual(data.compare, (0, 1))


if __name__ == "__main__":
    unittest.main()
