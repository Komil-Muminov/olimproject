import unittest
import sys
import os
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from catalog_manager import GibCatalogManager


class TestCatalogManager(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".json")
        self.tmp.close()
        self.mgr = GibCatalogManager(self.tmp.name)

    def tearDown(self):
        if os.path.exists(self.tmp.name):
            os.remove(self.tmp.name)

    def test_default_loaded(self):
        self.assertEqual(self.mgr.get("2111"), "Музди меҳнат (Заработная плата)")
        self.assertIn("2211", self.mgr.get_all())

    def test_add_and_update(self):
        self.mgr.set("9999", "Санҷиш")
        self.assertEqual(self.mgr.get("9999"), "Санҷиш")
        self.mgr.set("9999", "Санҷиши нав")
        self.assertEqual(self.mgr.get("9999"), "Санҷиши нав")

    def test_delete(self):
        self.mgr.set("8888", "Барои нест кардан")
        self.assertTrue(self.mgr.delete("8888"))
        self.assertIsNone(self.mgr.get("8888"))
        self.assertFalse(self.mgr.delete("8888"))

    def test_persistence(self):
        self.mgr.set("7777", "Махсус")
        # Создаем новый экземпляр из того же файла
        mgr2 = GibCatalogManager(self.tmp.name)
        self.assertEqual(mgr2.get("7777"), "Махсус")

    def test_reset_to_default(self):
        self.mgr.delete("2111")
        self.assertIsNone(self.mgr.get("2111"))
        self.mgr.reset_to_default()
        self.assertIsNotNone(self.mgr.get("2111"))


if __name__ == "__main__":
    unittest.main()
