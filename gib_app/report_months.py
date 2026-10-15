"""Лист «По месяцам»: сумма и число платежей реестров по месяцам и столбчатый график — сразу видны провалы
и всплески. Месяцы далеко от периода файлов (странные даты) тоже показаны — они маленькие."""
from openpyxl.cell import WriteOnlyCell
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font

from registry_stats import MONTHS
from report_fast import title_rows
from report_styles import BORDER, CENTER, NUM, fill_color

HEAD_ROW = 4
CHART_AT = "F4"


def write_months_sheet(wb, data):
    """Только если есть реестры с датами."""
    if not data.months:
        return None
    ws = wb.create_sheet("По месяцам")
    ws.sheet_properties.tabColor = "0891B2"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCD", (16, 20, 12, 3)):
        ws.column_dimensions[col].width = w
    title_rows(ws, "ПЛАТЕЖИ ПО МЕСЯЦАМ", "0891B2", "Сумма и число платежей всех выбранных реестров по месяцам "
               "(по дате документа). Провал или всплеск на графике — повод посмотреть этот месяц.", 3)
    head = []
    for t in ("Месяц", "Сумма", "Платежей"):
        c = WriteOnlyCell(ws, value=t)
        c.font, c.fill, c.border, c.alignment = Font(bold=True), fill_color("E2E8F0"), BORDER, CENTER
        head.append(c)
    ws.append(head)
    for (y, m), (total, count) in data.months.items():
        c = WriteOnlyCell(ws, value=round(total, 2))
        c.number_format = NUM
        ws.append([f"{MONTHS[m - 1]} {y}", c, count])
    last = HEAD_ROW + len(data.months)
    chart = BarChart()
    chart.title, chart.y_axis.title, chart.legend = "Сумма платежей по месяцам", "Сумма", None
    chart.height, chart.width = 9, 22
    chart.add_data(Reference(ws, min_col=2, min_row=HEAD_ROW, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=HEAD_ROW + 1, max_row=last))
    ws.add_chart(chart, CHART_AT)
    return ws
