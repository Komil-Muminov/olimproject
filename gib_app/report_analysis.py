"""Excel глубокого анализа: лист «Анализ» (цвет строки — что случилось, столбец «Ошибки», легенда и
подсказки при наведении) и лист «Ошибки подробно»."""
from collections import Counter
from datetime import datetime

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from analysis_model import (APPEARED, DETAIL_HEAD, FELL, FLAG_HINTS, FLAGS, GAP, GONE, GREW, SAME, STATUS_HINTS,
                            status)
from report_fast import table_sheet, title_rows
from report_styles import BORDER, CENTER, NUM, fill_color

COLORS = {APPEARED: "DCFCE7", GONE: "FEE2E2", GREW: "DBEAFE", FELL: "FEF9C3", SAME: "FFFFFF", GAP: "EDE9FE"}
ERR_COLOR = "FED7AA"
HEAD_ROW = 5
AUTHOR = "Распределение по ГИБ"


def _cf_fill(color: str):
    """Заливка для условного форматирования: Excel берёт цвет фона (bgColor), а не цвет узора."""
    return PatternFill("solid", start_color=color, end_color=color)


def _cell(ws, v, hint=None, **style):
    c = WriteOnlyCell(ws, value=v)
    for k, val in style.items():
        setattr(c, k, val)
    if hint:
        c.comment = Comment(hint, AUTHOR, width=320, height=110)
    return c


def write_analysis_report(path, data, progress=lambda p, t: None) -> dict:
    parts, with_role = data.parts, len(data.roles) > 1
    rows = sorted(((e, status(e, parts)) for e in data.entities), key=lambda r: -r[0].total)
    progress(15, "Пишу лист «Анализ»…")
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Анализ")
    ws.sheet_view.showGridLines = False
    head = ["№", "ИНН", "Название"] + (["Кто"] if with_role else []) + [p.label for p in parts] + \
        ["Всего", "Что случилось", "Ошибки", "Названий"]
    n = len(head)
    first_sum = 4 + with_role
    st_col, err_col = first_sum + len(parts) + 1, first_sum + len(parts) + 2
    for i, w in enumerate([7, 13, 46] + [12] * with_role + [17] * len(parts) + [18, 18, 30, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"D{HEAD_ROW + 1}"
    order = " → ".join(f"«{p.label}»" for p in parts)
    title_rows(ws, "ГЛУБОКИЙ АНАЛИЗ", "0F172A", f"Создано {datetime.now():%d.%m.%Y %H:%M} · части по порядку: {order}"
               + (". «Что случилось» — между первой и последней частью." if len(parts) > 1 else ""), n)
    _legend(ws, len(parts) > 1)
    hints = _head_hints(parts, with_role)
    ws.append([_cell(ws, t, hints.get(t), font=Font(bold=True), fill=fill_color("E2E8F0"), border=BORDER,
                     alignment=CENTER) for t in head])
    for k, (e, st) in enumerate(rows, 1):
        sums = [(_cell(ws, e.sums[i], number_format=NUM) if p.has_sums else "есть") if e.counts[i] else None
                for i, p in enumerate(parts)]
        ws.append([k, e.inn, e.name] + ([e.role] if with_role else []) + sums
                  + [_cell(ws, e.total, number_format=NUM), st, ", ".join(f for f in FLAGS if f in e.flags),
                     e.names if e.names > 1 else None])
        if k % 20000 == 0:
            progress(15 + int(60 * k / len(rows)), "Пишу лист «Анализ»…")
    last = HEAD_ROW + len(rows)
    if rows:
        letter = get_column_letter
        ws.auto_filter.ref = f"A{HEAD_ROW}:{letter(n)}{last}"
        for st, color in COLORS.items():
            if color != "FFFFFF":
                ws.conditional_formatting.add(f"A{HEAD_ROW + 1}:{letter(n)}{last}", FormulaRule(
                    formula=[f'${letter(st_col)}{HEAD_ROW + 1}="{st}"'], fill=_cf_fill(color)))
        ws.conditional_formatting.add(f"{letter(err_col)}{HEAD_ROW + 1}:{letter(err_col)}{last}", FormulaRule(
            formula=[f'LEN(${letter(err_col)}{HEAD_ROW + 1})>0'], fill=_cf_fill(ERR_COLOR), font=Font(bold=True)))
    progress(80, "Пишу лист «Ошибки подробно»…")
    table_sheet(wb, "Ошибки подробно", "ОШИБКИ ПОДРОБНО", "C2410C",
                "Каждая строка — одна найденная ошибка: где она, что не так и как, вероятно, должно быть.",
                ["№"] + DETAIL_HEAD, [7, 18, 16, 36, 22, 46, 12, 12, 16, 50],
                [[k] + r for k, r in enumerate(sorted(data.details, key=lambda r: FLAGS.index(r[0])), 1)],
                money=(9,), dates=(8,))
    progress(95, "Сохраняю…")
    wb.save(path)
    return {"path": path, "rows": len(rows), "statuses": Counter(st for _, st in rows if st),
            "flags": Counter(f for e, _ in rows for f in e.flags)}


def _legend(ws, compare: bool):
    """Строка 4: цвета «что случилось» и цвет ошибок — с подсказками при наведении."""
    cells = [None, None, _cell(ws, "Цвет строки (наведите — подсказка) →", font=Font(bold=True),
                               alignment=Alignment(horizontal="right", vertical="center"))]
    if compare:
        cells += [_cell(ws, st, STATUS_HINTS[st], fill=fill_color(c), border=BORDER, alignment=CENTER)
                  for st, c in COLORS.items()]
    else:
        cells += [_cell(ws, "одна часть — сравнивать не с чем", font=Font(italic=True, color="64748B"))]
    cells += [_cell(ws, "Ошибки", "\n".join(f"• {f}: {FLAG_HINTS[f]}" for f in FLAGS) +
                    "\nПодробно — на листе «Ошибки подробно».", fill=fill_color(ERR_COLOR), border=BORDER,
                    alignment=CENTER, font=Font(bold=True))]
    ws.append(cells)


def _head_hints(parts, with_role) -> dict:
    hints = {"ИНН": "ИНН организации (как записан в файле).",
             "Название": "Самое частое написание названия у этого ИНН.",
             "Кто": "плательщик — кто платил; получатель — кому платили; организация — из обычной таблицы с ИНН.",
             "Всего": "Сумма по всем частям.",
             "Что случилось": "Сравнение первой и последней части:\n" +
                              "\n".join(f"• {st}: {h}" for st, h in STATUS_HINTS.items()),
             "Ошибки": "\n".join(f"• {f}: {h}" for f, h in FLAG_HINTS.items()),
             "Названий": "Сколько разных написаний названия у этого ИНН (пусто — одно)."}
    hints.update({p.label: f"Сумма за часть «{p.label}» (файл {p.source})." if p.has_sums else
                  f"В таблице «{p.label}» (файл {p.source}) нет сумм — «есть» значит, что организация там есть."
                  for p in parts})
    return hints
