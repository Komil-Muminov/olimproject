"""Колонки от руки (как в Olimv3): разбор записи, настройка чтения листа, листы файла, шапка ГИБ."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gib_app")))

from openpyxl import Workbook  # noqa: E402

from column_entry import to_mapping  # noqa: E402
from column_spec import SpecError, parse  # noqa: E402
from file_sheets import read_sheets  # noqa: E402
from registry_columns import guess_header  # noqa: E402
from report_source import data_rows, head_row  # noqa: E402

INNS = [f"{100000000 + i * 7919:09d}" for i in range(3)]


def _book(sheets: dict) -> str:
    path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    wb.save(path)
    return path


class ColumnEntryTest(unittest.TestCase):
    def test_parse(self):
        self.assertEqual(parse("c"), [(2, None, None)])
        self.assertEqual(parse("C5"), [(2, 5, None)])
        self.assertEqual(parse("c5-300"), [(2, 5, 300)])
        self.assertEqual(parse("B5-B20, С5-С20"), [(1, 5, 20), (2, 5, 20)])     # «С» — русская буква
        for bad in ("5C", "C5-D9", "C9-C5", "C5-"):
            with self.assertRaises(SpecError):
                parse(bad)

    def test_rows_from_inn_field(self):
        """«A4-A5» — только строки 4–5; «A» — с первой строки, где ИНН (подписи — из настоящей шапки, не из «ИТОГО»)."""
        path = _book({"Л": [["Свод"], ["ИНН", "Название", "Сумма"], [None, "ИТОГО", 60],
                            [INNS[0], "А", 10], [INNS[1], "Б", 20], [INNS[2], "В", 30]]})
        sheet = read_sheets(path)["sheets"]["Л"]
        m = to_mapping("Л", sheet, {"inn": "a4-a5", "sums": "c"})
        self.assertEqual((m["head"], m["last"], m["sums"]), (2, 5, [2]))
        self.assertEqual([r for r, _ in data_rows(path, m)], [4, 5])
        m = to_mapping("Л", sheet, {"inn": "a", "sums": "c"})
        self.assertEqual((m["head"], m["labels_row"]), (2, 1))
        self.assertEqual(head_row(path, m), ["ИНН", "Название", "Сумма"])

    def test_errors(self):
        sheet = read_sheets(_book({"Л": [["ИНН"], ["нет"]]}))["sheets"]["Л"]
        for typed in ({}, {"inn": "a, b"}, {"inn": "a"}):          # нет колонки, две колонки, ИНН не найден
            with self.assertRaises(SpecError):
                to_mapping("Л", sheet, typed)

    def test_table_not_from_a1(self):
        """Таблица с колонки C: столбцы считаются от первой заполненной колонки листа."""
        path = _book({"Л": [[None, None, "ИНН", "Сумма"], [None, None, INNS[0], 5]]})
        sheet = read_sheets(path)["sheets"]["Л"]
        m = to_mapping("Л", sheet, {"inn": "c", "sums": "d"})
        self.assertEqual((sheet["col0"], m["cols"]["inn"], m["sums"]), (2, 0, [1]))
        with self.assertRaises(SpecError):
            to_mapping("Л", sheet, {"inn": "c", "sums": "a"})        # колонка A на этом листе пустая

    def test_sheets_and_empty(self):
        p = read_sheets(_book({"Янв": [["ИНН"], [INNS[0]]], "Пусто": [], "Фев": [["ИНН"], [INNS[1]]]}))
        self.assertEqual(list(p["sheets"]), ["Янв", "Фев"])           # пустой лист пропущен

    def test_guess_header(self):
        """Шапка для «Заполнить ГИБ»: строка, где узнано больше всего заголовков."""
        g = guess_header([["Отчёт"], [None], ["Колонка1", "Сумма", "Колонка3"], [1, 100, "x"]])
        self.assertEqual((g["head"], g["cols"]["amount"], g["heads"][0]), (2, 1, "Колонка1"))


if __name__ == "__main__":
    unittest.main()
