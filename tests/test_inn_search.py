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
        self.assertEqual(rows[0][:4], ("Файл", "Лист", "Строка", "30 фоиза.xlsx: A РМА"))
        self.assertEqual([r[:3] for r in rows[1:]], [("30 фоиза.xlsx", "Л", 2), ("M1-4.xlsx", "Л", 2), ("M1-4.xlsx", "Л", 4)])
        self.assertIn("дивиденд", rows[-1])

    def test_not_found(self):
        a = _save("a.xlsx", [["ИНН"], ["630000962"]])
        self.assertIsNone(write_inn_report(os.path.join(tempfile.mkdtemp(), "r.xlsx"), [a], {a: _map(a, inn="a")}, INN)["path"])


if __name__ == "__main__":
    unittest.main()
