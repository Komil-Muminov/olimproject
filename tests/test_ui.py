import unittest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from PySide6.QtWidgets import QApplication


class TestMainWindow(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

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

    def test_analysis_screen_filters(self):
        from ui import MainWindow
        from analysis_lists import ListCollector
        from analysis_scan import build
        from registry_stats import Registry
        import ui_analysis_flow
        rows = [["Номи корхона", "РМА", "Маблағ 2024", None, "Номи корхона", "РМА", "Маблағ 2025"]]
        rows += [[f"Орг {i}", f"0100000{i:02d}", 100.0, None, f"Орг {i}", f"0100000{i:02d}", 100.0 + 7.5 * i]
                 for i in range(12)]
        c = ListCollector()
        c.add_file("t.xlsx", {"Лист1": rows})
        w = MainWindow()
        ui_analysis_flow._scanned(w, {"data": build(Registry(), c, ["t.xlsx"], {})})
        self.assertEqual((w.an_parts.count(), w.an_model.rowCount()), (2, 12))
        w.an_status["вырос"].setChecked(True)                 # фильтр «вырос»: 11 из 12
        self.assertEqual(w.an_model.rowCount(), 11)
        self.assertIn("вырос (11)", w.an_status["вырос"].text())
        w.an_query.setText("010000005")
        ui_analysis_flow.apply(w)
        self.assertEqual(w.an_model.rowCount(), 1)
        w.close()

if __name__ == "__main__":
    unittest.main()
