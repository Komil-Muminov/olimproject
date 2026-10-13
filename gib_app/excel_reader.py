"""Чтение Excel файлов через openpyxl и базовые утилиты адресации ячеек."""
import os
import re
import tempfile
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

from constants import XL_OPENXML

CELL_RE = re.compile(r"^\s*([A-Za-z]{1,3})\s*(\d+)\s*$")
GIB_HEADER_RE = re.compile(r"^\s*(\d{3,4})(?!\d)")


def parse_cell(addr: str):
    """'B5' -> (5, 2). None, если адрес некорректный."""
    m = CELL_RE.match(addr or "")
    if not m:
        return None
    return int(m.group(2)), column_index_from_string(m.group(1).upper())


def cell_name(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def gib_code(value) -> str | None:
    """Заголовок столбца -> код ГИБ (2111, '2111', '2111 Музди маош')."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    m = GIB_HEADER_RE.match(str(value))
    return m.group(1) if m else None


def _xls_to_xlsx(path: str) -> str:
    """Старый .xls читаем через Excel: сохраняем временную копию в .xlsx."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    out = os.path.join(tempfile.gettempdir(), f"gib_preview_{os.getpid()}.xlsx")
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
        wb.SaveAs(out, FileFormat=XL_OPENXML)
        wb.Close(False)
    finally:
        xl.Quit()
    return out


class Book:
    """Значения всех листов книги для показа в интерфейсе."""

    def __init__(self, path: str):
        self.path = path
        read_path = _xls_to_xlsx(path) if path.lower().endswith(".xls") else path
        wb = openpyxl.load_workbook(read_path, read_only=True, data_only=True)
        self.sheets: dict[str, list[list]] = {}
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            width = max((len(r) for r in rows), default=0)
            while rows and all(v is None for v in rows[-1]):
                rows.pop()
            self.sheets[ws.title] = [r + [None] * (width - len(r)) for r in rows]
        wb.close()

    @staticmethod
    def value(data, row: int, col: int):
        if 1 <= row <= len(data) and 1 <= col <= len(data[row - 1]):
            return data[row - 1][col - 1]
        return None
