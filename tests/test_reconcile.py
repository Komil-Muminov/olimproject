import os
import sys
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from analysis_model import ORG, PAYER, RECIPIENT, AnalysisData, Entity, Part  # noqa: E402
from analysis_reconcile import NOT_PAID, PAID, reconcile  # noqa: E402


class ReconcileTest(unittest.TestCase):
    def test_list_against_registry(self):
        parts = [Part("Список", "30.xlsx", "list", (2024, 0, 0)), Part("Янв–Апр 2024", "M1-4.xlsx", "payments", (2024, 1, 0))]
        ents = [Entity("40058284", "А (список)", ORG, [0.0, 0.0], [1, 0]),          # ноль впереди потерян
                Entity("630000962", "Б (список)", ORG, [0.0, 0.0], [1, 0]),
                Entity("040058284", "А", PAYER, [0.0, 500.0], [0, 2]),
                Entity("630000962", "Б", RECIPIENT, [0.0, 7.0], [0, 1]),
                Entity("111111111", "В", PAYER, [0.0, 1.0], [0, 1])]               # не из списка — не в сверке
        reg, rows = reconcile(AnalysisData(parts, ents, {}, [], [], None, "files"))
        self.assertEqual([p.label for p in reg], ["Янв–Апр 2024"])
        self.assertEqual([(r.name, r.paid, r.payments, r.got, r.result) for r in rows],
                         [("Б (список)", [0.0], 0, 7.0, NOT_PAID), ("А (список)", [500.0], 2, 0.0, PAID)])

    def test_who_hint(self):
        from report_reconcile import TOP, who_hint
        entries = [(f"Получатель {k}", f"{k:09d}", [float(k), 0.0], [1, 0]) for k in range(1, TOP + 3)]
        text = who_hint(entries, "Кому платил:", 0)
        lines = text.split(chr(10))
        self.assertEqual(lines[1], f"• Получатель {TOP + 2} ({TOP + 2:09d}): {TOP + 2},00 — 1 плат.")   # крупные сверху
        self.assertEqual(lines[-1], "…и ещё 2 на 3,00")
        self.assertIsNone(who_hint(entries, "Кому платил:", 1))          # во второй части платежей нет

    def test_nothing_to_reconcile(self):
        parts = [Part("Янв", "M1.xlsx", "payments", (2024, 1, 0))]
        self.assertEqual(reconcile(AnalysisData(parts, [], {}, [], [], None, "files")), ([], []))


if __name__ == "__main__":
    unittest.main()
