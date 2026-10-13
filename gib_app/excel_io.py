"""Фасад для чтения и записи Excel файлов."""

from constants import (
    NO_GIB_NOTE, RED_FILL, RED_FONT, SUM_ROW_LABELS, TOTAL_HEADERS,
    XL_CALC_MANUAL, XL_DOWN, XL_FORMAT_FROM_LEFT_OR_ABOVE, XL_GREEN,
    XL_NO_GIB, XL_NO_GIB_FONT, XL_OPENXML, XL_TO_RIGHT, XL_XLSM, XL_YELLOW
)
from excel_automation import embed_automation
from excel_com import WriteError
from excel_reader import Book, cell_name, gib_code, parse_cell
from excel_writer import fill_in_place

__all__ = [
    "Book",
    "WriteError",
    "fill_in_place",
    "embed_automation",
    "parse_cell",
    "cell_name",
    "gib_code",
    "XL_TO_RIGHT",
    "XL_DOWN",
    "XL_FORMAT_FROM_LEFT_OR_ABOVE",
    "XL_CALC_MANUAL",
    "XL_OPENXML",
    "XL_XLSM",
    "XL_GREEN",
    "XL_YELLOW",
    "XL_NO_GIB",
    "XL_NO_GIB_FONT",
    "NO_GIB_NOTE",
    "TOTAL_HEADERS",
    "SUM_ROW_LABELS",
    "RED_FILL",
    "RED_FONT",
]
