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

    def test_registry_inn_search(self):
        from ui import MainWindow
        from registry_stats import Registry, Party
        w = MainWindow()
        reg = Registry()
        f = reg.start_file("test.xlsx", "Лист1")
        p = Party("040000011")
        p.names["ООО Тест"] = 1
        p.sums[f] = 50000.0
        p.counts[f] = 3
        reg.payers["040000011"] = p
        reg.finish()
        w.registry = reg
        self.assertTrue(w.reg_found_box.isHidden())
        w.reg_inn.setText(" 40000011 ")                     # потерянный ноль и пробелы
        w.reg_find_btn.click()
        self.assertIn("040000011", w.reg_found.text())
        self.assertFalse(w.reg_found_box.isHidden())
        self.assertFalse(w.reg_export_btn.isHidden())
        w.reg_inn.setText("123")
        w.reg_find_btn.click()
        self.assertIn("не найден", w.reg_found.text())
        self.assertTrue(w.reg_export_btn.isHidden())
        w.close()

if __name__ == "__main__":
    unittest.main()
