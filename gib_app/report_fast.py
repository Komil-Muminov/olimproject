"""Быстрая запись больших таблиц (сотни тысяч строк) в режиме write_only."""
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font

from report_styles import fill_color


def _cell(ws, v, **style):
    c = WriteOnlyCell(ws, value=v)
    for k, val in style.items():
        setattr(c, k, val)
    return c


def title_rows(ws, title: str, color: str, hint: str, n: int):
    """Строка 1 — цветной заголовок на n столбцов, строка 2 — пояснение, строка 3 — пустая."""
    band = fill_color(color)
    ws.append([_cell(ws, title, font=Font(bold=True, size=14, color="FFFFFF"), fill=band,
                     alignment=Alignment(vertical="center", indent=1))]
              + [_cell(ws, None, fill=band) for _ in range(n - 1)])
    ws.row_dimensions[1].height = 30
    ws.append([_cell(ws, hint, font=Font(italic=True, color="475569"))])
    ws.append([])
