"""Быстрая запись больших таблиц (сотни тысяч строк) в режиме write_only."""
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from report_styles import BORDER, CENTER, NUM, fill_color

DATE = "dd.mm.yyyy"
HEAD_ROW = 4


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


def table_sheet(wb, name, title, color, hint, head, widths, rows, money=(), dates=()):
    """Лист-таблица: заголовок с количеством, пояснение, шапка с фильтром, закреплённая шапка.
    money / dates — номера столбцов (с 1) с суммами / датами."""
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = color
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"A{HEAD_ROW + 1}"
    n = len(head)
    title_rows(ws, f"{title} — {len(rows)}", color, hint, n)
    head_style = dict(font=Font(bold=True, color="0F172A"), fill=fill_color("E2E8F0"), border=BORDER, alignment=CENTER)
    ws.append([_cell(ws, t, **head_style) for t in head])
    ws.row_dimensions[HEAD_ROW].height = 32
    if not rows:
        ws.append([None, _cell(ws, "нет ✓", font=Font(italic=True, color="16A34A"))])
        return ws
    money, dates = {c - 1 for c in money}, {c - 1 for c in dates}
    for row in rows:
        ws.append([_cell(ws, v, number_format=NUM) if i in money and v is not None
                   else _cell(ws, v, number_format=DATE) if i in dates and v is not None
                   else v for i, v in enumerate(row)])
    ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(n)}{HEAD_ROW + len(rows)}"
    return ws


def text_cell(ws, v, bold=False, color="334155", size=11, link=None, fmt=None):
    """Ячейка для листа «Итого» (обычный текст / ссылка на лист / число)."""
    c = _cell(ws, v, font=Font(bold=bold, size=size, color=color, underline="single" if link else None))
    if link:
        c.hyperlink = link
    if fmt:
        c.number_format = fmt
    return c
