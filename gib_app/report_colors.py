"""Цвета «что случилось» и ошибок — общие для листов Excel и экрана результата."""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from analysis_model import APPEARED, FELL, GAP, GONE, GREW, SAME

COLORS = {APPEARED: "DCFCE7", GONE: "FEE2E2", GREW: "DBEAFE", FELL: "FEF9C3", SAME: "FFFFFF", GAP: "EDE9FE"}
ERR_COLOR = "FED7AA"


def cf_fill(color: str):
    """Заливка для условного форматирования: Excel берёт цвет фона (bgColor), а не цвет узора."""
    return PatternFill("solid", start_color=color, end_color=color)


def add_row_colors(ws, first_row: int, last_row: int, n_cols: int, status_col: int, error_col: int = None):
    """Цвет всей строки — по столбцу «Что случилось»; ячейка «Ошибки» — оранжевая, если не пустая."""
    if last_row < first_row:
        return
    letter = get_column_letter
    for st, color in COLORS.items():
        if color != "FFFFFF":
            ws.conditional_formatting.add(f"A{first_row}:{letter(n_cols)}{last_row}", FormulaRule(
                formula=[f'${letter(status_col)}{first_row}="{st}"'], fill=cf_fill(color)))
    if error_col:
        ws.conditional_formatting.add(f"{letter(error_col)}{first_row}:{letter(error_col)}{last_row}", FormulaRule(
            formula=[f'LEN(${letter(error_col)}{first_row})>0'], fill=cf_fill(ERR_COLOR), font=Font(bold=True)))
