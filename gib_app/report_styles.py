"""Стили и оформление для генерации Excel-отчетов."""
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NUM = "#,##0.00"
THIN = Side(style="thin", color="CBD5E1")
MED = Side(style="medium", color="64748B")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
VCENTER = Alignment(vertical="center")

# категория: заголовок, цвет, светлый цвет
COLORS = {
    "new": ("CA8A04", "FEF9C3"),
    "gone": ("DC2626", "FEE2E2"),
    "changed": ("2563EB", "DBEAFE"),
    "same": ("16A34A", "DCFCE7"),
    "both": ("16A34A", "DCFCE7"),
    "dup": ("7C3AED", "EDE9FE"),
    "names": ("0891B2", "CFFAFE"),
}


def fill_color(color: str):
    return PatternFill("solid", fgColor=color)


def styled_cell(ws, r: int, c: int, v=None, **style):
    x = ws.cell(r, c)
    if v is not None:
        x.value = v
    for k, val in style.items():
        setattr(x, k, val)
    return x


def apply_diff_colors(ws, col: int, r1: int, r2: int):
    """Разница: рост зелёным, уменьшение красным."""
    rng = f"{get_column_letter(col)}{r1}:{get_column_letter(col)}{r2}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0.004"],
                                                  font=Font(color="15803D", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["-0.004"],
                                                  font=Font(color="B91C1C", bold=True)))
