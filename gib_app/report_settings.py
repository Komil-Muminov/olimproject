"""Лист «Как считали»: когда сделан анализ, какие файлы и листы взяты, какие столбцы выбраны (ИНН, суммы,
особенности), что скрыто в файлах. Чтобы через месяц было видно, откуда взялись цифры."""
import os
from datetime import datetime

from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font

from analysis_periods import MODE_TITLES
from report_fast import title_rows
from report_styles import BORDER, CENTER, fill_color

KINDS = {"registry": "реестр платежей", "list": "список организаций"}
OPTIONS = (("transpose", "таблица перевёрнута (организации по столбцам)"), ("by_name", "ИНН нет — по названию"),
           ("repeats", "организация в нескольких строках — не ошибка"), ("hidden_sheet", "скрытый лист"))


def _head(ws, texts):
    cells = []
    for t in texts:
        c = WriteOnlyCell(ws, value=t)
        c.font, c.fill, c.border, c.alignment = Font(bold=True), fill_color("E2E8F0"), BORDER, CENTER
        cells.append(c)
    ws.append(cells)


def _sheet_rows(fname: str, m: dict) -> list:
    """Одна настройка листа → строки: что выбрано и где."""
    out = [[fname, m["sheet"], KINDS.get(m["kind"], m["kind"]), "Строка заголовков", m["row0"] + m["head"] + 1]]
    out += [[fname, m["sheet"], "", title, col] for title, col in m.get("labels", {}).items()]
    sums = m.get("sum_labels") or []
    out.append([fname, m["sheet"], "", "Суммы", "; ".join(sums) if sums else "не выбраны (только «есть / нет»)"])
    if len(sums) > 1:
        out.append([fname, m["sheet"], "", "Как суммы", "каждый столбец отдельно" if m.get("separate") else "сложены"])
    opts = [text for key, text in OPTIONS if m.get(key)]
    if opts:
        out.append([fname, m["sheet"], "", "Особенности", "; ".join(opts)])
    return out


def write_settings_sheet(wb, data, mappings: dict = None):
    ws = wb.create_sheet("Как считали")
    ws.sheet_properties.tabColor = "64748B"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", (34, 26, 20, 22, 70)):
        ws.column_dimensions[col].width = w
    title_rows(ws, "КАК СЧИТАЛИ", "64748B", "Параметры анализа: что взято из каких файлов. Исходные файлы не менялись.", 5)
    a, b = data.compare or (None, None)
    for k, v in (("Дата анализа", datetime.now().strftime("%d.%m.%Y %H:%M")), ("Файлов", len(data.paths)),
                 ("Части", MODE_TITLES.get(data.mode, data.mode)),
                 ("Сравнение", f"«{data.parts[a].label}» → «{data.parts[b].label}»" if data.compare else "нет")):
        ws.append([WriteOnlyCell(ws, value=k), v])
    ws.append([])
    _head(ws, ["Файл", "Лист", "Что это", "Что выбрано", "Столбец / значение"])
    for path in data.paths:
        fname, chosen = os.path.basename(path), (mappings or {}).get(path)
        if isinstance(chosen, list):
            for row in (r for m in chosen for r in _sheet_rows(fname, m)):
                ws.append(row)
        else:
            ws.append([fname, "", "", "Столбцы", "определены программой автоматически"])
    if data.notes:
        ws.append([])
        _head(ws, ["Файл", "Лист", "Скрытое", "", "Что это значит"])
        for n in data.notes:
            what = ("скрытый лист; " if n["hidden_sheet"] else "") + (f"скрыто строк: {n['hidden_rows']}" if n["hidden_rows"] else "")
            hint = ("каждая скрытая строка — на листах «Скрытые строки»" if n["listed"]
                    else "скрыта почти вся таблица (например, был фильтр) — строки по отдельности не выписаны")
            ws.append([os.path.basename(n["file"]), n["sheet"], what.strip("; "), "", hint if n["hidden_rows"] else ""])
    return ws
