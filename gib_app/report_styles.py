"""Стили и оформление для генерации Excel-отчетов."""
from openpyxl.styles import Alignment, Border, PatternFill, Side

NUM = "#,##0.00"
THIN = Side(style="thin", color="CBD5E1")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def fill_color(color: str):
    return PatternFill("solid", fgColor=color)
