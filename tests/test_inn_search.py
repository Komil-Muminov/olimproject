"""«Один ИНН во всех файлах»: все строки ИНН из всех файлов (колонки — из карточек), без предела по строкам."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook, load_workbook  # noqa: E402

from analysis_propose import propose  # noqa: E402
from column_entry import to_mapping  # noqa: E402
from report_inn_search import clean_inn, write_inn_report  # noqa: E402

INN = "040058284"


def _save(name, rows):
    path = os.path.join(tempfile.mkdtemp(), name)
    wb = Workbook()
    wb.active.title = "Л"
    for r in rows:
        wb.active.append(r)
    wb.save(path)
    return path


def _map(path, **typed):
    return [to_mapping("Л", propose(path)["sheets"]["Л"], typed, {})]


class InnSearchTest(unittest.TestCase):
    def test_clean(self):
        self.assertEqual([clean_inn(" 40058284 "), clean_inn("040 058 284"), clean_inn("abc")], [INN, INN, ""])

    def test_rows_from_all_files(self):
        a = _save("30 фоиза.xlsx", [["РМА", "Номи", "30 фоиза"], [INN, "Шабака", 39], ["630000962", "Фурудгоҳ", 27]])
        b = _save("M1-4.xlsx", [["Дата", "ИНН", "Сумма", "Основание"], ["09.01.2024", INN, 100, "налог"],
                                ["10.01.2024", "630000962", 5, "x"], ["11.01.2024", int(INN), 50, "дивиденд"]])
        maps = {a: _map(a, inn="a", name="b", sums="c"), b: _map(b, inn="b", sums="c")}
        out = os.path.join(tempfile.mkdtemp(), "r.xlsx")
        res = write_inn_report(out, [a, b], maps, "40058284")              # ноль впереди потерян — не страшно
        self.assertEqual((res["rows"], res["files"]), (3, [("30 фоиза.xlsx", 1, 39.0), ("M1-4.xlsx", 2, 150.0)]))
        rows = list(load_workbook(out).active.iter_rows(min_row=8, values_only=True))
        self.assertEqual(rows[0][:5], ("Файл", "Лист", "Строка", "Сумма строки", "30 фоиза.xlsx: A РМА"))
        self.assertEqual([r[:3] for r in rows[1:]], [("30 фоиза.xlsx", "Л", 2), ("M1-4.xlsx", "Л", 2), ("M1-4.xlsx", "Л", 4)])
        self.assertIn("дивиденд", rows[-1])
        res = write_inn_report(os.path.join(tempfile.mkdtemp(), "p.xlsx"), [a, b], maps, INN, threshold=(60, None))
        self.assertEqual(res["files"], [("M1-4.xlsx", 1, 100.0)])           # порог «от 60»: 39 и 50 — мимо

    def test_threshold_sheet_in_analysis(self):
        """Порог «от 40» в «Анализе»: лист «Суммы от 40,00» — записи из всех файлов с суммой строки от 40."""
        from analysis_jobs import make_report
        from analysis_scan import auto_mode, scan_files
        a = _save("a.xlsx", [["ИНН", "Сумма"], [INN, 39], ["630000962", 45]])
        b = _save("b.xlsx", [["ИНН", "Сумма"], [INN, 100], ["630000962", 5]])
        maps = {a: _map(a, inn="a", sums="b"), b: _map(b, inn="a", sums="b")}
        raw = scan_files([a, b], maps)["raw"]
        out = os.path.join(tempfile.mkdtemp(), "r.xlsx")
        res = make_report(out, raw, auto_mode(raw), None, maps, (40, None))
        self.assertEqual(res["threshold"], {"text": "от 40,00", "rows": 2, "total": 145.0, "cut": False})
        ws = load_workbook(out)["Суммы от 40,00"]
        rows = [r[:4] for r in ws.iter_rows(min_row=9, values_only=True)]
        self.assertEqual(rows, [("a.xlsx", "Л", 3, 45), ("b.xlsx", "Л", 2, 100)])

    def test_inn_in_any_column(self):
        """ИНН не в колонке ИНН, а в основании — строка тоже находится."""
        a = _save("a.xlsx", [["ИНН", "Сумма", "Основание"], ["630000962", 7, f"за ООО «Шабака» (ИНН {INN})"],
                             ["630000962", 9, "налог 1040058284"]])               # 10 цифр — другой номер
        out = os.path.join(tempfile.mkdtemp(), "r.xlsx")
        res = write_inn_report(out, [a], {a: _map(a, inn="a", sums="b")}, INN)
        self.assertEqual(res["files"], [("a.xlsx", 1, 7.0)])

    def test_not_found(self):
        a = _save("a.xlsx", [["ИНН"], ["630000962"]])
        self.assertIsNone(write_inn_report(os.path.join(tempfile.mkdtemp(), "r.xlsx"), [a], {a: _map(a, inn="a")}, INN)["path"])


if __name__ == "__main__":
    unittest.main()
