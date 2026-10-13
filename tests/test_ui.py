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

    def test_registry_inn_selector(self):
        from ui import MainWindow
        from registry_index import clean_inn_query
        from ui_registry_inn import populate_inn_view
        from registry_stats import Registry, Party
        w = MainWindow()
        self.assertTrue(hasattr(w, "reg_table"))
        self.assertEqual(w.reg_table.columnCount(), 4)

        self.assertEqual(clean_inn_query("040000011  ·  ООО Тест (100 ₽)"), "040000011")
        self.assertEqual(clean_inn_query("  040000011  "), "040000011")

        reg = Registry()
        f = reg.start_file("test.xlsx", "Лист1")
        p = Party("040000011")
        p.names["ООО Тест"] = 1
        p.sums[f] = 50000.0
        p.counts[f] = 3
        reg.payers["040000011"] = p
        reg.finish()

        populate_inn_view(w, reg)
        self.assertEqual(w.reg_table.rowCount(), 1)
        self.assertEqual(w.reg_table.item(0, 0).text(), "040000011")

        w.registry = reg
        w.reg_table.cellClicked.emit(0, 0)
        self.assertIn("040000011", w.reg_found.text())
        self.assertFalse(w.reg_export_btn.isHidden())
        w.close()


if __name__ == "__main__":
    unittest.main()
