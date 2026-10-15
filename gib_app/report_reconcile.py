"""Лист «Сверка со списком»: организации из списков — платили ли они по реестрам (только если выбраны
и списки, и реестры)."""
from openpyxl.cell import WriteOnlyCell
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from analysis_reconcile import PAID, reconcile
from report_fast import title_rows
from report_styles import BORDER, CENTER, NUM, fill_color

HEAD_ROW = 4
COLOR = {PAID: "DCFCE7", "": "FEE2E2"}
HINTS = {"Где в списке": "В каких таблицах (листах) списка есть эта организация.",
         "Заплатил всего": "Сумма платежей, где эта организация — плательщик (по всем реестрам).",
         "Платежей": "Сколько платежей она сделала.",
         "Получил": "Сумма платежей, где эта организация — получатель.",
         "Итог": "Платил — есть хотя бы один платёж в реестрах. Не платил — в реестрах её нет среди плательщиков."}


def _cell(ws, v, **style):
    c = WriteOnlyCell(ws, value=v)
    for k, val in style.items():
        setattr(c, k, val)
    return c


def write_reconcile_sheet(wb, data) -> dict | None:
    """→ {"paid", "not_paid"} или None, если сверять нечего."""
    reg, rows = reconcile(data)
    if not rows:
        return None
    ws = wb.create_sheet("Сверка со списком")
    ws.sheet_properties.tabColor = "16A34A"
    ws.sheet_view.showGridLines = False
    head = ["№", "ИНН", "Название", "Где в списке"] + [f"Заплатил: {p.label}" for p in reg] + \
        ["Заплатил всего", "Платежей", "Получил", "Итог"]
    for i, w in enumerate([6, 13, 46, 34] + [18] * len(reg) + [18, 10, 16, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"D{HEAD_ROW + 1}"
    paid = sum(r.result == PAID for r in rows)
    title_rows(ws, f"СВЕРКА СО СПИСКОМ — платили {paid} из {len(rows)}", "16A34A",
               "Каждая организация из списка (таблицы с ИНН): есть ли её платежи в реестрах. Не платившие — сверху.",
               len(head))
    hints = dict(HINTS, **{f"Заплатил: {p.label}": f"Заплатил за «{p.label}» ({p.source})." for p in reg})
    cells = []
    for t in head:
        c = _cell(ws, t, font=Font(bold=True), fill=fill_color("E2E8F0"), border=BORDER, alignment=CENTER)
        if t in hints:
            c.comment = Comment(hints[t], "Распределение по ГИБ", width=320, height=90)
        cells.append(c)
    ws.append(cells)
    money = lambda v: _cell(ws, v, number_format=NUM) if v else None
    for k, r in enumerate(rows, 1):
        ws.append([k, r.inn, r.name, r.where] + [money(v) for v in r.paid]
                  + [money(round(sum(r.paid), 2)), r.payments or None, money(r.got),
                     _cell(ws, r.result, fill=fill_color(COLOR.get(r.result, COLOR[""])), font=Font(bold=True))])
    ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(len(head))}{HEAD_ROW + len(rows)}"
    return {"paid": paid, "not_paid": len(rows) - paid}
