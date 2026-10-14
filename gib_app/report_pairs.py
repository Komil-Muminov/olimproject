"""Лист «Кто кому платит»: все пары плательщик → получатель из реестров — суммы по частям, изменение
«было → стало», цвет строки — что случилось с этой связкой."""
from openpyxl.cell import WriteOnlyCell
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from analysis_model import change, status, status_hints
from report_colors import add_row_colors
from report_fast import title_rows
from report_styles import BORDER, CENTER, NUM, fill_color

HEAD_ROW = 4
HINTS = {"Плательщик": "Кто платил (самое частое написание названия).",
         "Получатель": "Кому платили.",
         "Всего": "Сумма всех платежей этого плательщика этому получателю.",
         "Платежей": "Сколько было платежей."}


def _cell(ws, v, **style):
    c = WriteOnlyCell(ws, value=v)
    for k, val in style.items():
        setattr(c, k, val)
    return c


def write_pairs_sheet(wb, data):
    """Только для реестров (у обычных таблиц нет пар). Крупные связки сверху."""
    if not data.pairs:
        return None
    parts, pair = data.parts, data.compare
    ws = wb.create_sheet("Кто кому платит")
    ws.sheet_properties.tabColor = "2563EB"
    ws.sheet_view.showGridLines = False
    change_head = ["Изменение", "Изменение, %"] if pair else []
    head = ["№", "Плательщик", "ИНН плательщика", "Получатель", "ИНН получателя"] + [p.label for p in parts] + \
        ["Всего", "Платежей"] + change_head + ["Что случилось"]
    widths = [7, 40, 13, 40, 13] + [17] * len(parts) + [18, 10] + [18, 12] * bool(pair) + [18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"C{HEAD_ROW + 1}"
    what = f" Цвет строки — что случилось со связкой: «{parts[pair[0]].label}» → «{parts[pair[1]].label}»." if pair else ""
    title_rows(ws, f"КТО КОМУ ПЛАТИТ — {len(data.pairs):,}".replace(",", " "), "2563EB",
               "Каждая строка — плательщик → получатель: сколько он ему заплатил в каждой части." + what, len(head))
    hints = dict(HINTS, **{"Что случилось": "\n".join(status_hints(parts, pair).values())})
    hints.update({p.label: f"Сумма платежей этой связки за «{p.label}»." for p in parts})
    row = []
    for t in head:
        c = _cell(ws, t, font=Font(bold=True), fill=fill_color("E2E8F0"), border=BORDER, alignment=CENTER)
        if hints.get(t):
            c.comment = Comment(hints[t], "Распределение по ГИБ", width=320, height=100)
        row.append(c)
    ws.append(row)
    rows = sorted(data.pairs, key=lambda x: -sum(x.sums))
    for k, x in enumerate(rows, 1):
        extra = []
        if pair:
            delta, pct = change(x, parts, pair)
            extra = [_cell(ws, delta, number_format=NUM) if delta is not None else None,
                     _cell(ws, pct, number_format="0.0%") if pct is not None else None]
        ws.append([k, x.payer, x.payer_inn, x.rec, x.rec_inn]
                  + [_cell(ws, x.sums[i], number_format=NUM) if x.counts[i] else None for i in range(len(parts))]
                  + [_cell(ws, round(sum(x.sums), 2), number_format=NUM), sum(x.counts)] + extra
                  + [status(x, parts, pair)])
    last = HEAD_ROW + len(rows)
    ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(len(head))}{last}"
    add_row_colors(ws, HEAD_ROW + 1, last, len(head), len(head))
    return ws
