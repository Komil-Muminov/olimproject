import unittest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from planner import Plan


class TestPlanner(unittest.TestCase):
    def test_plan_loading(self):
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        sample_path = os.path.join(root, "BANK 2.xlsx")
        if os.path.exists(sample_path):
            plan = Plan(sample_path)
            self.assertIsNotNone(plan.sheet)


if __name__ == "__main__":
    unittest.main()
