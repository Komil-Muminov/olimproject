"""Файл с результатом сравнения по ИНН — один лист «Различия».

Разделы друг под другом (Новые, Пропавшие, Изменения, Дубликаты, Без изменений), у каждого
цветная полоса-заголовок, своя таблица и подытог; раздел можно свернуть. Внизу — «ИТОГО»
по категориям и всего по спискам. Исходный файл не меняется.
"""
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from compare import sorted_changes

NUM = "#,##0.00"
THIN = Side(style="thin", color="CBD5E1")
MED = Side(style="medium", color="64748B")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# раздел: заголовок, пояснение, цвет полосы, светлый цвет подытога/итога
SECTIONS = {
    "new": ("🆕 НОВЫЕ", "есть только во 2-м списке", "CA8A04", "FEF9C3"),
    "gone": ("❌ ПРОПАВШИЕ", "были только в 1-м списке", "DC2626", "FEE2E2"),
    "changed": ("🔄 ИЗМЕНИЛАСЬ СУММА", "есть в обоих списках, сумма другая", "2563EB", "DBEAFE"),
    "dup": ("👥 ДУБЛИКАТЫ", "один ИНН записан в списке несколько раз", "7C3AED", "EDE9FE"),
    "names": ("✏ РАЗНЫЕ НАЗВАНИЯ", "один ИНН, а название в списках записано по-разному", "0891B2", "CFFAFE"),
    "same": ("✓ БЕЗ ИЗМЕНЕНИЙ", "есть в обоих списках, суммы совпадают", "16A34A", "DCFCE7"),
    "both": ("✓ ЕСТЬ В ОБОИХ СПИСКАХ", "суммы не сравнивались — во 2-м списке их нет", "16A34A", "DCFCE7"),
}


def report_path(src, sheet):
    """«файл — различия (лист).xlsx» рядом с исходным; если уже есть — «… (2).xlsx»."""
    base = f"{os.path.splitext(src)[0]} — различия ({sheet})"
    path, n = base + ".xlsx", 2
    while os.path.exists(path):
        path, n = f"{base} ({n}).xlsx", n + 1
    return path


def _fill(color):
    return PatternFill("solid", fgColor=color)


def write_compare_report(path, src, sheet, diff, a, b):
    y1, y2 = a.year or "Левая таблица", b.year or "Правая таблица"
    in1 = f"в {a.year}" if a.year else "в левой таблице"
    in2 = f"в {b.year}" if b.year else "в правой таблице"
    hints = {
        "new": f"есть {in2}, не было {in1}",
        "gone": f"были {in1}, нет {in2}",
        "changed": f"есть и {in1}, и {in2} — сумма другая",
        "dup": "один ИНН записан в списке несколько раз",
        "same": f"есть и {in1}, и {in2} — суммы совпадают",
        "both": f"есть и {in1}, и {in2} (суммы не сравнивались — {in2} их нет)",
        "names": f"один ИНН, а название {in1} и {in2} записано по-разному",
    }
    metrics = [p[0] for p in diff.pairs]
    wb = Workbook()
    ws = wb.active
    ws.title = "Различия"
    ws.sheet_properties.outlinePr.summaryBelow = False      # кнопка сворачивания — у полосы раздела
    ws.sheet_view.showGridLines = False

    # столбцы: № | ИНН | Организация | (было | стало | разница) × показатель | Примечание
    first_val = 4
    n_cols = 3 + 3 * len(metrics) + 1
    note_col = n_cols
    for i, w in enumerate([6, 13, 48] + [16, 16, 15] * len(metrics) + [44], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def cell(r, c, v=None, **style):
        x = ws.cell(r, c)
        if v is not None:
            x.value = v
        for k, val in style.items():
            setattr(x, k, val)
        return x

    def band(r, text, color, font_color="FFFFFF", size=12, height=26):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
        cell(r, 1, text, font=Font(bold=True, size=size, color=font_color), fill=_fill(color),
             alignment=Alignment(vertical="center", indent=1))
        for c in range(2, n_cols + 1):
            ws.cell(r, c).fill = _fill(color)
        ws.row_dimensions[r].height = height

    # шапка листа
    cell(1, 1, "Сравнение двух списков по ИНН / РМА", font=Font(bold=True, size=16, color="0F172A"))
    ws.row_dimensions[1].height = 28
    cell(2, 1, f"Файл: {os.path.basename(src)}   ·   лист «{sheet}»   ·   создано {datetime.now():%d.%m.%Y %H:%M}",
         font=Font(color="64748B"))
    cell(3, 1, f"{a.year or 'Левая таблица'}: {len(a.rows)} организаций   →   "
               f"{b.year or 'правая таблица'}: {len(b.rows)} организаций",
         font=Font(bold=True, color="334155"))
    notes = []
    if diff.names_differ:
        notes.append(f"Сопоставлено по ИНН: у {diff.names_differ} организаций название записано по-разному — "
                     f"список в разделе «Разные названия».")
    cell(4, 1, "   ".join(notes) or None, font=Font(italic=True, color="64748B"))

    # шапка таблицы (две строки)
    h1, h2 = 6, 7
    for c, title in ((1, "№"), (2, "ИНН / РМА"), (3, "Организация"), (note_col, "Примечание")):
        ws.merge_cells(start_row=h1, start_column=c, end_row=h2, end_column=c)
        cell(h1, c, title)
    for i, m in enumerate(metrics):
        c = first_val + 3 * i
        ws.merge_cells(start_row=h1, start_column=c, end_row=h1, end_column=c + 2)
        cell(h1, c, m)
        for j, t in enumerate((y1, y2, "Разница")):
            cell(h2, c + j, t)
    for r in (h1, h2):
        for c in range(1, n_cols + 1):
            cell(r, c, font=Font(bold=True, color="0F172A"), fill=_fill("E2E8F0"), border=BORDER, alignment=CENTER)
    ws.row_dimensions[h1].height = 34
    ws.freeze_panes = ws.cell(h2 + 1, 4)

    row = h2 + 2
    subtotal_rows, counts, ranges = {}, {}, {}
    diff_cols = [first_val + 3 * i + 2 for i in range(len(metrics))]

    def section(key, items, note=lambda it: "", hint=None):
        """items: (ИНН, название, [(было, стало)] по показателям)."""
        nonlocal row
        title, _, color, light = SECTIONS[key]
        hint = hint or hints[key]
        counts[key] = len(items)
        band(row, f"{title} — {len(items)}   ·   {hint}", color)
        row += 1
        start = row
        if not items:
            cell(row, 3, "нет", font=Font(italic=True, color="94A3B8"))
            row += 1
        for i, (inn, name, vals) in enumerate(items, 1):
            cell(row, 1, i, alignment=Alignment(horizontal="center"))
            cell(row, 2, inn)
            cell(row, 3, name)
            for m, (v1, v2) in enumerate(vals):
                c = first_val + 3 * m
                cell(row, c, v1, number_format=NUM)
                cell(row, c + 1, v2, number_format=NUM)
                L1, L2 = get_column_letter(c), get_column_letter(c + 1)
                cell(row, c + 2, f"={L2}{row}-{L1}{row}", number_format=NUM)
            cell(row, note_col, note((inn, name, vals)) or None, font=Font(color="64748B"))
            for c in range(1, n_cols + 1):
                ws.cell(row, c).border = BORDER
            row += 1
        end = row - 1
        if metrics and items:
            cell(row, 3, "Итого по разделу", font=Font(bold=True))
            for c in range(first_val, first_val + 3 * len(metrics)):
                L = get_column_letter(c)
                cell(row, c, f"=SUM({L}{start}:{L}{end})", number_format=NUM, font=Font(bold=True))
            for c in range(1, n_cols + 1):
                cell(row, c, fill=_fill(light), border=BORDER)
            subtotal_rows[key] = row
            row += 1
        ranges[key] = (start, row - 1)
        for r in range(start, row):
            ws.row_dimensions[r].outlineLevel = 1
        row += 1

    def side_vals(vals, side):
        """значения одного списка → [(было, стало)]"""
        return [(v, None) if side == 1 else (None, v) for v in vals]

    by_size = lambda o: -sum(abs(v or 0) for pair in o[2] for v in pair)      # крупные суммы сверху
    section("new", sorted([(k, n, side_vals(v, 2)) for k, n, v in diff.org_only2], key=by_size))
    section("gone", sorted([(k, n, side_vals(v, 1)) for k, n, v in diff.org_only1], key=by_size))
    collapsed = []
    if metrics:
        section("changed", sorted_changes(diff))
    dup_notes = {(k, n): f"список {i} · строки {', '.join(map(str, rr))} ({len(rr)} раза)"
                 for i, k, n, rr in diff.duplicates}
    section("dup", [(k, n, [(None, None)] * len(metrics)) for _, k, n, _ in diff.duplicates],
            lambda it: dup_notes.get((it[0], it[1]), ""))
    if metrics:
        section("same", diff.org_same)
        collapsed.append("same")
    else:
        section("both", [(k, n, []) for k, n, _ in diff.org_same + diff.org_changed])
        collapsed.append("both")
    names_notes = {k: f"{in2}: {n2}" for k, _, n2 in diff.org_names}
    section("names", sorted([(k, n1, [(None, None)] * len(metrics)) for k, n1, _ in diff.org_names],
                            key=lambda o: o[1]),
            lambda it: names_notes.get(it[0], ""),
            hint=f"один ИНН: слева — как записано {in1}, справа в «Примечании» — как {in2}")
    # «Без изменений» / «Есть в обоих» свёрнуты: обычно это неинтересно (раскрыть — «+» слева)
    for key in collapsed:
        s, e = ranges[key]
        for r in range(s, e + 1):
            ws.row_dimensions[r].hidden = True

    # цвет разницы: рост зелёным, уменьшение красным
    for c in diff_cols:
        rng = f"{get_column_letter(c)}{h2 + 1}:{get_column_letter(c)}{row}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0.004"],
                                                      font=Font(color="15803D", bold=True)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["-0.004"],
                                                      font=Font(color="B91C1C", bold=True)))

    # ИТОГО
    row += 1
    band(row, "ИТОГО", "0F172A", size=14, height=32)
    row += 1
    if metrics:
        for r in (row, row + 1):
            for c in range(1, n_cols + 1):
                cell(r, c, font=Font(bold=True, color="0F172A"), fill=_fill("E2E8F0"), border=BORDER, alignment=CENTER)
        ws.merge_cells(start_row=row, start_column=1, end_row=row + 1, end_column=3)
        cell(row, 1, "Категория")
        for i, m in enumerate(metrics):
            c = first_val + 3 * i
            ws.merge_cells(start_row=row, start_column=c, end_row=row, end_column=c + 2)
            cell(row, c, m)
            for j, t in enumerate((y1, y2, "Разница")):
                cell(row + 1, c + j, t)
        row += 2
    top = row
    for key in [k for k in ("new", "gone", "changed", "dup", "same", "both") if k in counts]:
        title, _, color, light = SECTIONS[key]
        hint = hints[key]
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        cell(row, 1, f"{title}   —   {counts[key]}", font=Font(bold=True, size=12, color=color),
             alignment=Alignment(vertical="center", indent=1))
        if key in subtotal_rows:
            for c in range(first_val, first_val + 3 * len(metrics)):
                cell(row, c, f"={get_column_letter(c)}{subtotal_rows[key]}", number_format=NUM,
                     font=Font(bold=True, size=11), alignment=Alignment(vertical="center"))
        cell(row, note_col, hint, font=Font(color="64748B"), alignment=Alignment(vertical="center"))
        for c in range(1, n_cols + 1):
            cell(row, c, fill=_fill(light), border=BORDER)
        ws.row_dimensions[row].height = 24
        row += 1
    # всего по спискам
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    cell(row, 1, f"ВСЕГО ПО СПИСКАМ   {len(a.rows)} → {len(b.rows)} организаций",
         font=Font(bold=True, size=13, color="FFFFFF"), alignment=Alignment(vertical="center", indent=1))
    for c in range(first_val, first_val + 3 * len(metrics)):
        L = get_column_letter(c)
        cell(row, c, f"=SUM({L}{top}:{L}{row - 1})", number_format=NUM, font=Font(bold=True, size=12, color="FFFFFF"),
             alignment=Alignment(vertical="center"))
    cell(row, note_col, f"{y1} → {y2}", font=Font(bold=True, color="E2E8F0"), alignment=Alignment(vertical="center"))
    for c in range(1, n_cols + 1):
        cell(row, c, fill=_fill("334155"), border=Border(top=MED, bottom=MED, left=THIN, right=THIN))
    ws.row_dimensions[row].height = 30
    if metrics:
        row += 1
        cell(row, 3, "Проверка — суммы самих списков:", font=Font(italic=True, color="64748B"))
        for i, (_, s1, s2) in enumerate(diff.totals):
            c = first_val + 3 * i
            cell(row, c, s1, number_format=NUM, font=Font(italic=True, color="64748B"))
            cell(row, c + 1, s2, number_format=NUM, font=Font(italic=True, color="64748B"))
    if diff.org_names:
        row += 2
        cell(row, 1, f"✏ Разные названия при одном ИНН: {len(diff.org_names)} — см. раздел выше "
                     f"(на суммы не влияет, организации сопоставлены по ИНН).", font=Font(color="0891B2", bold=True))

    ws.print_title_rows = f"{h1}:{h2}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)
    return path
