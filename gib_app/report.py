"""Файл с результатом сравнения по ИНН: лист «Итого» и по листу на каждую категорию.

Листы: Итого · Новые (Только справа) · Пропавшие (Только слева) · Изменения · Без изменений
(Есть в обоих) · Дубликаты · Разные названия. На «Итого» — что с чем сравнивали и цветная
сводка; названия категорий там — ссылки на их листы. Исходный файл не меняется.
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
VCENTER = Alignment(vertical="center")

# категория: заголовок, цвет, светлый цвет
COLORS = {
    "new": ("CA8A04", "FEF9C3"),
    "gone": ("DC2626", "FEE2E2"),
    "changed": ("2563EB", "DBEAFE"),
    "same": ("16A34A", "DCFCE7"),
    "both": ("16A34A", "DCFCE7"),
    "dup": ("7C3AED", "EDE9FE"),
    "names": ("0891B2", "CFFAFE"),
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


def _cell(ws, r, c, v=None, **style):
    x = ws.cell(r, c)
    if v is not None:
        x.value = v
    for k, val in style.items():
        setattr(x, k, val)
    return x


def _diff_colors(ws, col, r1, r2):
    """Разница: рост зелёным, уменьшение красным."""
    rng = f"{get_column_letter(col)}{r1}:{get_column_letter(col)}{r2}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0.004"],
                                                  font=Font(color="15803D", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["-0.004"],
                                                  font=Font(color="B91C1C", bold=True)))


def _category_sheet(wb, sheet_name, title, hint, key, items, metrics, y1, y2,
                    name_title="Организация", note_title="Примечание", note=lambda it: ""):
    """Лист категории: цветной заголовок, пояснение, таблица, «Итого», фильтр.
    items: (ИНН, название, [(было, стало)] по показателям). → (лист, строка «Итого» или None)."""
    color, light = COLORS[key]
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = color
    ws.sheet_view.showGridLines = False
    first_val = 4
    n_cols = 3 + 3 * len(metrics) + 1
    note_col = n_cols
    for i, w in enumerate([6, 13, 50] + [16, 16, 15] * len(metrics) + [50], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    _cell(ws, 1, 1, f"{title} — {len(items)}", font=Font(bold=True, size=14, color="FFFFFF"),
          fill=_fill(color), alignment=Alignment(vertical="center", indent=1))
    for c in range(2, n_cols + 1):
        ws.cell(1, c).fill = _fill(color)
    ws.row_dimensions[1].height = 30
    _cell(ws, 2, 1, hint, font=Font(italic=True, color="475569"))

    h1, h2 = 4, 5
    for c, t in ((1, "№"), (2, "ИНН / РМА"), (3, name_title), (note_col, note_title)):
        ws.merge_cells(start_row=h1, start_column=c, end_row=h2, end_column=c)
        _cell(ws, h1, c, t)
    for i, m in enumerate(metrics):
        c = first_val + 3 * i
        ws.merge_cells(start_row=h1, start_column=c, end_row=h1, end_column=c + 2)
        _cell(ws, h1, c, m)
        for j, t in enumerate((y1, y2, "Разница")):
            _cell(ws, h2, c + j, t)
    for r in (h1, h2):
        for c in range(1, n_cols + 1):
            _cell(ws, r, c, font=Font(bold=True, color="0F172A"), fill=_fill("E2E8F0"), border=BORDER,
                  alignment=CENTER)
    ws.row_dimensions[h1].height = 34
    ws.freeze_panes = ws.cell(h2 + 1, 4)

    row = h2 + 1
    if not items:
        _cell(ws, row, 3, "нет", font=Font(italic=True, color="94A3B8"))
        return ws, None
    for i, (inn, name, vals) in enumerate(items, 1):
        _cell(ws, row, 1, i, alignment=Alignment(horizontal="center"))
        _cell(ws, row, 2, inn)
        _cell(ws, row, 3, name)
        for m, (v1, v2) in enumerate(vals):
            c = first_val + 3 * m
            _cell(ws, row, c, v1, number_format=NUM)
            _cell(ws, row, c + 1, v2, number_format=NUM)
            L1, L2 = get_column_letter(c), get_column_letter(c + 1)
            _cell(ws, row, c + 2, f"={L2}{row}-{L1}{row}", number_format=NUM)
        _cell(ws, row, note_col, note((inn, name, vals)) or None, font=Font(color="475569"))
        for c in range(1, n_cols + 1):
            ws.cell(row, c).border = BORDER
        row += 1
    last = row - 1
    ws.auto_filter.ref = f"A{h2}:{get_column_letter(n_cols)}{last}"
    for i in range(len(metrics)):
        _diff_colors(ws, first_val + 3 * i + 2, h2 + 1, last + 1)
    if not metrics:
        return ws, None
    _cell(ws, row, 3, "Итого", font=Font(bold=True))
    for c in range(first_val, first_val + 3 * len(metrics)):
        L = get_column_letter(c)
        _cell(ws, row, c, f"=SUM({L}{h2 + 1}:{L}{last})", number_format=NUM, font=Font(bold=True))
    for c in range(1, n_cols + 1):
        _cell(ws, row, c, fill=_fill(light), border=Border(top=MED, bottom=THIN, left=THIN, right=THIN))
    ws.print_title_rows = f"{h1}:{h2}"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return ws, row


def write_compare_report(path, src, sheet, diff, a, b):
    # годы — если они есть у обеих частей, иначе «левая / правая часть» (как видно в исходном листе)
    years = bool(a.year and b.year)
    y1, y2 = (a.year, b.year) if years else (f"Левая ({a.span})", f"Правая ({b.span})")
    in1 = f"в {a.year}" if years else f"в левой части ({a.span})"
    in2 = f"в {b.year}" if years else f"в правой части ({b.span})"
    metrics = [p[0] for p in diff.pairs]

    # (ключ, имя листа, заголовок, пояснение)
    cats = [
        ("new", "Новые" if years else "Только справа",
         "🆕 НОВЫЕ" if years else f"🆕 ТОЛЬКО В ПРАВОЙ ЧАСТИ ({b.span})",
         f"Организации, которые есть {in2}, но которых не было {in1}." if years
         else f"Организации, которые есть {in2}, но которых нет {in1}."),
        ("gone", "Пропавшие" if years else "Только слева",
         "❌ ПРОПАВШИЕ" if years else f"❌ ТОЛЬКО В ЛЕВОЙ ЧАСТИ ({a.span})",
         f"Организации, которые были {in1}, но которых нет {in2}." if years
         else f"Организации, которые есть {in1}, но которых нет {in2}."),
    ]
    if metrics:
        cats += [("changed", "Изменения", "🔄 ИЗМЕНИЛАСЬ СУММА",
                  f"Организации, которые есть и {in1}, и {in2}, но сумма у них другая. Крупные разницы — сверху."),
                 ("same", "Без изменений", "✓ БЕЗ ИЗМЕНЕНИЙ",
                  f"Организации, которые есть и {in1}, и {in2}, и суммы совпадают.")]
    else:
        cats += [("both", "Есть в обоих", "✓ ЕСТЬ В ОБОИХ СПИСКАХ",
                  f"Организации, которые есть и {in1}, и {in2}. Суммы не сравнивались — {in2} их нет.")]
    cats += [("dup", "Дубликаты", "👥 ДУБЛИКАТЫ",
              "Один и тот же ИНН записан в списке несколько раз — проверьте, не задвоена ли организация."),
             ("names", "Разные названия", "✏ РАЗНЫЕ НАЗВАНИЯ",
              f"Один ИНН, а название записано по-разному. Организации всё равно сопоставлены по ИНН.")]

    def side_vals(vals, side):
        return [(v, None) if side == 1 else (None, v) for v in vals]

    by_size = lambda o: -sum(abs(v or 0) for pair in o[2] for v in pair)      # крупные суммы сверху
    dup_notes = {(k, n): f"{'левая' if i == 1 else 'правая'} часть · строки {', '.join(map(str, rr))} "
                         f"({len(rr)} раза)" for i, k, n, rr in diff.duplicates}
    names2 = {k: n2 for k, _, n2 in diff.org_names}
    data = {
        "new": (sorted([(k, n, side_vals(v, 2)) for k, n, v in diff.org_only2], key=by_size), {}),
        "gone": (sorted([(k, n, side_vals(v, 1)) for k, n, v in diff.org_only1], key=by_size), {}),
        "changed": (sorted_changes(diff), {}),
        "same": (diff.org_same, {}),
        "both": ([(k, n, []) for k, n, _ in diff.org_same + diff.org_changed], {}),
        "dup": ([(k, n, [(None, None)] * len(metrics)) for _, k, n, _ in diff.duplicates],
                dict(note_title="Где повторяется", note=lambda it: dup_notes.get((it[0], it[1]), ""))),
        "names": (sorted([(k, n1, [(None, None)] * len(metrics)) for k, n1, _ in diff.org_names], key=lambda o: o[1]),
                  dict(name_title=f"Как записано {in1}", note_title=f"Как записано {in2}",
                       note=lambda it: names2.get(it[0], ""))),
    }

    wb = Workbook()
    summary = wb.active
    summary.title = "Итого"
    subtotals = {}
    for key, sheet_name, title, hint in cats:
        items, extra = data[key]
        tbl_metrics = metrics if key in ("new", "gone", "changed", "same") else []
        _, sub = _category_sheet(wb, sheet_name, title, hint, key, items, tbl_metrics, y1, y2, **extra)
        subtotals[key] = (sheet_name, sub, len(items))

    # ---------- лист «Итого»
    ws = summary
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "0F172A"
    first_val = 3
    n_cols = 2 + 3 * len(metrics) + 1
    note_col = n_cols
    for i, w in enumerate([44, 14] + [16, 16, 15] * len(metrics) + [70], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    _cell(ws, 1, 1, "Сравнение двух списков по ИНН / РМА", font=Font(bold=True, size=16, color="0F172A"))
    ws.row_dimensions[1].height = 28
    _cell(ws, 2, 1, f"Файл: {os.path.basename(src)}   ·   лист «{sheet}»   ·   создано {datetime.now():%d.%m.%Y %H:%M}",
          font=Font(color="64748B"))
    _cell(ws, 3, 1, f"Сравниваются: {a.where()}   ↔   {b.where()}", font=Font(bold=True, color="334155"))
    _cell(ws, 4, 1, "Каждая категория — на своём листе (вкладки внизу). Название категории в таблице — ссылка на её лист.",
          font=Font(italic=True, color="64748B"))

    row = 6
    head = ["Категория", "Организаций"]
    for r in (row, row + 1):
        for c in range(1, n_cols + 1):
            _cell(ws, r, c, font=Font(bold=True, color="0F172A"), fill=_fill("E2E8F0"), border=BORDER,
                  alignment=CENTER)
    for c, t in enumerate(head, 1):
        ws.merge_cells(start_row=row, start_column=c, end_row=row + 1, end_column=c)
        _cell(ws, row, c, t)
    ws.merge_cells(start_row=row, start_column=note_col, end_row=row + 1, end_column=note_col)
    _cell(ws, row, note_col, "Что это значит")
    for i, m in enumerate(metrics):
        c = first_val + 3 * i
        ws.merge_cells(start_row=row, start_column=c, end_row=row, end_column=c + 2)
        _cell(ws, row, c, m)
        for j, t in enumerate((y1, y2, "Разница")):
            _cell(ws, row + 1, c + j, t)
    ws.row_dimensions[row].height = 34
    row += 2

    top = row
    total_keys = [k for k, *_ in cats if k not in ("dup", "names")]
    for key, sheet_name, title, hint in cats:
        if key == "names":
            continue
        color, light = COLORS[key]
        name, sub, count = subtotals[key]
        link = _cell(ws, row, 1, title, font=Font(bold=True, size=12, color=color, underline="single"),
                     alignment=Alignment(vertical="center", indent=1))
        link.hyperlink = f"#'{name}'!A1"
        _cell(ws, row, 2, count, font=Font(bold=True, size=12), alignment=Alignment(horizontal="center", vertical="center"))
        if sub and key in total_keys:
            for c in range(len(metrics) * 3):
                _cell(ws, row, first_val + c, f"='{name}'!{get_column_letter(4 + c)}{sub}", number_format=NUM,
                      font=Font(bold=True), alignment=VCENTER)
        _cell(ws, row, note_col, hint, font=Font(color="475569"), alignment=VCENTER)
        for c in range(1, n_cols + 1):
            _cell(ws, row, c, fill=_fill(light), border=BORDER)
        ws.row_dimensions[row].height = 26
        row += 1
    # всего по спискам (без дубликатов — это не отдельные организации)
    _cell(ws, row, 1, "ВСЕГО ПО СПИСКАМ", font=Font(bold=True, size=13, color="FFFFFF"),
          alignment=Alignment(vertical="center", indent=1))
    _cell(ws, row, 2, f"{len(a.rows)} → {len(b.rows)}", font=Font(bold=True, size=12, color="FFFFFF"),
          alignment=Alignment(horizontal="center", vertical="center"))
    for c in range(len(metrics) * 3):
        L = get_column_letter(first_val + c)
        rows = [str(top + i) for i, (k, *_) in enumerate(c_ for c_ in cats if c_[0] != "names") if k in total_keys]
        _cell(ws, row, first_val + c, "=" + "+".join(f"{L}{r}" for r in rows), number_format=NUM,
              font=Font(bold=True, size=12, color="FFFFFF"), alignment=VCENTER)
    _cell(ws, row, note_col, f"{a.where()}  →  {b.where()}", font=Font(bold=True, color="E2E8F0"), alignment=VCENTER)
    for c in range(1, n_cols + 1):
        _cell(ws, row, c, fill=_fill("334155"), border=Border(top=MED, bottom=MED, left=THIN, right=THIN))
    ws.row_dimensions[row].height = 30
    if metrics:
        row += 1
        _cell(ws, row, 1, "Проверка — суммы самих списков:", font=Font(italic=True, color="64748B"),
              alignment=Alignment(indent=1))
        for i, (_, s1, s2) in enumerate(diff.totals):
            c = first_val + 3 * i
            _cell(ws, row, c, s1, number_format=NUM, font=Font(italic=True, color="64748B"))
            _cell(ws, row, c + 1, s2, number_format=NUM, font=Font(italic=True, color="64748B"))
    for i in range(len(metrics)):
        _diff_colors(ws, first_val + 3 * i + 2, top, row)
    # разные названия — отдельной строкой: на суммы не влияет
    row += 2
    name, _, count = subtotals["names"]
    link = _cell(ws, row, 1, f"✏ РАЗНЫЕ НАЗВАНИЯ — {count}", font=Font(bold=True, color=COLORS["names"][0],
                                                                       underline="single"))
    link.hyperlink = f"#'{name}'!A1"
    _cell(ws, row, 2, None)
    _cell(ws, row, 3, "один ИНН, а название записано по-разному — на суммы не влияет", font=Font(color="475569"))

    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    wb.save(path)
    return path
