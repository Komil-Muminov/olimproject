"""Excel-отчёт «Свод по ИНН» по реестрам платежей: итоги, плательщики, получатели, кто пропал /
появился, резкий рост и падение, ошибки."""
import os
from datetime import datetime

from openpyxl import Workbook

from registry_analysis import JUMP, gone_and_new, jumps, party_rows
from registry_stats import MONTHS
from report_fast import table_sheet, text_cell, title_rows
from report_registry_errors import write_error_sheets
from report_styles import NUM

SHEETS = [
    ("Плательщики", "0F766E", "каждый ИНН плательщика одной строкой: суммы по периодам и за всё время"),
    ("Получатели", "2563EB", "то же для получателей"),
    ("Пропали", "DC2626", "платили (получали) раньше, а в последнем периоде — нет"),
    ("Появились", "CA8A04", "в первом периоде не было, потом появились"),
    ("Рост и падение", "7C3AED", f"сумма к прошлому периоду выросла или упала в {JUMP:g} раза и больше"),
    ("Двойные платежи", "B91C1C", "одинаковые №, дата, ИНН плательщика, ИНН получателя и сумма"),
    ("Неправильные ИНН", "EA580C", "ИНН не из 9 цифр: «0», потерянный ноль, буквы"),
    ("Странные даты", "64748B", "дата документа не из периода файла (2004, 2029…)"),
]


def _party_table(wb, name, rows, labels, roles=None):
    """Таблица ИНН; roles — «плательщик»/«получатель» для каждой строки (листы «Пропали», «Появились»)."""
    _, color, hint = next(x for x in SHEETS if x[0] == name)
    who = ["Кто"] if roles else []
    head = ["№"] + who + ["ИНН", "Название"] + labels + ["Всего", "Платежей", "Где встречается"]
    widths = [7] + [12] * len(who) + [13, 46] + [17] * len(labels) + [18, 10, 30]
    data = [[i] + ([roles[i - 1]] if roles else []) + [r.inn, r.name] + r.sums + [r.total, sum(r.counts), r.status]
            for i, r in enumerate(rows, 1)]
    first = 4 + len(who)
    table_sheet(wb, name, name.upper(), color, hint, head, widths, data, money=range(first, first + len(labels) + 1))


def _with_roles(payers: list, recs: list, key) -> tuple:
    both = sorted([(r, "плательщик") for r in payers] + [(r, "получатель") for r in recs], key=key)
    return [r for r, _ in both], [w for _, w in both]


def write_registry_report(path: str, reg, progress=lambda p, t: None) -> dict:
    """→ {"path", "counts": [(лист, сколько строк)]} — списком, чтобы порядок листов сохранился."""
    labels = [reg.files[f].label for f in reg.order]
    progress(10, "Считаю плательщиков и получателей…")
    payers, recs = party_rows(reg.payers, reg), party_rows(reg.recs, reg)
    (g1, n1), (g2, n2) = gone_and_new(payers), gone_and_new(recs)
    gone = _with_roles(g1, g2, key=lambda x: -x[0].total)
    new = _with_roles(n1, n2, key=lambda x: -x[0].total)
    jj = _with_roles(jumps(payers, reg), jumps(recs, reg), key=lambda x: -abs(x[0][5]))
    counts = {"Плательщики": len(payers), "Получатели": len(recs), "Пропали": len(gone[0]),
              "Появились": len(new[0]), "Рост и падение": len(jj[0]), "Двойные платежи": len(reg.dups),
              "Неправильные ИНН": len(reg.bad_inn), "Странные даты": len(reg.odd_dates)}
    wb = Workbook(write_only=True)
    _summary(wb, reg, counts)
    progress(25, "Пишу плательщиков и получателей…")
    _party_table(wb, "Плательщики", payers, labels)
    _party_table(wb, "Получатели", recs, labels)
    progress(55, "Пишу, кто пропал и кто появился…")
    _party_table(wb, "Пропали", gone[0], labels, gone[1])
    _party_table(wb, "Появились", new[0], labels, new[1])
    progress(70, "Пишу рост и падение…")
    table_sheet(wb, "Рост и падение", "РЕЗКИЙ РОСТ И ПАДЕНИЕ", SHEETS[4][1],
                SHEETS[4][2] + ". Самые большие изменения сверху.",
                ["№", "Кто", "ИНН", "Название", "С периода", "На период", "Было", "Стало", "Разница", "Во сколько раз"],
                [7, 12, 13, 46, 15, 15, 17, 17, 17, 12],
                [[i, who, r.inn, r.name, p1, p2, a, b, d, round(b / a, 1) if b >= a else round(-a / b, 1)]
                 for i, ((r, p1, p2, a, b, d), who) in enumerate(zip(*jj), 1)], money=(7, 8, 9))
    progress(80, "Пишу ошибки…")
    write_error_sheets(wb, reg, SHEETS)
    progress(92, "Сохраняю файл…")
    wb.save(path)
    return {"path": path, "counts": list(counts.items())}


def _summary(wb, reg, counts: dict):
    ws = wb.create_sheet("Итого")
    ws.sheet_properties.tabColor = "0F172A"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCD", (34, 16, 22, 70)):
        ws.column_dimensions[col].width = w
    title_rows(ws, "Свод по ИНН — реестры платежей", "0F172A", f"Создано {datetime.now():%d.%m.%Y %H:%M}. "
               "Исходные файлы не менялись.", 4)
    ws.append([text_cell(ws, t, bold=True) for t in ("Файл", "Период", "Платежей", "Сумма")])
    for f in reg.order:
        info = reg.files[f]
        ws.append([os.path.basename(info.path), info.label, info.rows, text_cell(ws, round(info.total, 2), fmt=NUM)])
    ws.append([text_cell(ws, "Всего", bold=True), "", sum(i.rows for i in reg.files),
               text_cell(ws, round(sum(i.total for i in reg.files), 2), bold=True, fmt=NUM)])
    ws.append([])
    ws.append([text_cell(ws, t, bold=True) for t in ("Лист (нажмите, чтобы открыть)", "Найдено", "", "Что это")])
    for name, color, hint in SHEETS:
        ws.append([text_cell(ws, name, bold=True, color=color, link=f"#'{name}'!A1"), counts[name], "", hint])
    ws.append([])
    ws.append([text_cell(ws, t, bold=True) for t in ("Месяц", "Платежей", "Сумма")])
    months = {}
    for info in reg.files:
        for m in info.main_months:
            c, s = months.get(m, (0, 0.0))
            months[m] = (c + info.months[m], s + info.month_sums[m])
    for (y, m), (c, s) in sorted(months.items()):
        ws.append([f"{MONTHS[m - 1]} {y}", c, text_cell(ws, round(s, 2), fmt=NUM)])
