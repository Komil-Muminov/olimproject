import unittest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from PySide6.QtWidgets import QApplication


class TestMainWindow(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_worker_returns_list(self):
        """Фоновая задача отдаёт не только словарь: просмотр файлов перед анализом возвращает список."""
        from worker import Worker
        got = []
        w = Worker(lambda progress: [{"path": "a.xlsx"}])
        w.done.connect(got.append)
        w.run()                                 # в этом же потоке — сигнал доходит сразу
        self.assertEqual(got, [[{"path": "a.xlsx"}]])

    def test_window_starts(self):
        from ui import MainWindow
        w = MainWindow()
        self.assertEqual(w.pages.count(), 4)
        self.assertIn("Справочник", w.catalog_btn.text())
        self.assertTrue(w.fill_box.isVisibleTo(w))          # ГИБ → «Заполнить ГИБ»: два файла
        w.embed_radio.setChecked(True)
        self.assertFalse(w.fill_box.isVisibleTo(w))
        w.mode_group.button(1).click()                      # Глубокий анализ: один файл
        self.assertFalse(w.gib_box.isVisibleTo(w))
        self.assertTrue(w.drop.isVisibleTo(w))
        w.close()

    def test_catalog_dialog_opens(self):
        from catalog_dialog import CatalogDialog
        dlg = CatalogDialog()
        self.assertGreater(dlg.table.rowCount(), 0)
        dlg.close()

    def test_help_dialog(self):
        """Кнопка «Как пользоваться» в шапке; окно — про оба режима и листы результата."""
        from ui import MainWindow
        from ui_help import HelpDialog
        w = MainWindow()
        self.assertIn("Как пользоваться", w.help_btn.text())
        text = HelpDialog(w).text.toPlainText()
        for part in ("Заполнить ГИБ", "своя карточка", "C5-C300", "Одна орг. во всех файлах", "Ошибки"):
            self.assertIn(part, text)
        w.close()

    def test_analysis_screen(self):
        """По карточке на файл; «Анализ» нажимается, когда у всех взятых файлов указана колонка ИНН."""
        import tempfile
        import time
        from openpyxl import Workbook
        from ui import MainWindow
        import ui_analysis_flow
        paths = []
        for name in ("M1-4.xlsx", "M5-8.xlsx"):
            paths.append(os.path.join(tempfile.mkdtemp(), name))
            wb = Workbook()
            wb.active.append(["ИНН", "Сумма"])
            wb.active.append(["040000011", 5])
            wb.save(paths[-1])
        w = MainWindow()
        ui_analysis_flow.start(w, paths)
        while w.worker:
            self.app.processEvents()
            time.sleep(0.02)
        self.app.processEvents()
        self.assertEqual(w.pages.currentIndex(), ui_analysis_flow.ANALYSIS_PAGE)
        self.assertEqual([os.path.basename(c.path) for c in w.an_file_cards], ["M1-4.xlsx", "M5-8.xlsx"])
        self.assertFalse(w.an_run_btn.isEnabled())               # колонки ещё не вписаны
        w.an_file_cards[0].edits["inn"].setText("a")
        self.assertFalse(w.an_run_btn.isEnabled())               # второй файл не заполнен
        w.an_file_cards[1].use.setChecked(False)                 # …и не берётся
        self.assertTrue(w.an_run_btn.isEnabled())
        w.close()

if __name__ == "__main__":
    unittest.main()
