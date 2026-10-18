"""Excel глубокого анализа: лист «Анализ» (суммы по частям, в каких файлах есть организация, столбец «Ошибки»,
подсказки при наведении) и листы с ошибками подробно."""
from collections import Counter
from datetime import datetime

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from analysis_model import FLAG_HINTS, FLAGS, present_files
from report_colors import ERR_COLOR, add_error_color, cf_fill
from report_errors import write_errors_sheets
from report_months import write_months_sheet
from report_all_by_inn import write_all_by_inn_sheet
from report_raw_sheets import write_raw_sheets
from report_reconcile import write_reconcile_sheet
from report_settings import write_settings_sheet
from report_pairs import write_pairs_sheet
from report_fast import title_rows
from report_styles import BORDER, CENTER, NUM, fill_color

HEAD_ROW = 5
AUTHOR = "Распределение по ГИБ"


def _name_cell(ws, e):
    """Название; если у ИНН несколько названий — при наведении все варианты (первые 20)."""
    if len(e.variants) < 2:
        return e.name
    more = f"\n…и ещё {len(e.variants) - 20}" if len(e.variants) > 20 else ""
    return _cell(ws, e.name, "Как записано в файлах:\n" + "\n".join(f"• {v}" for v in e.variants[:20]) + more)


def _cell(ws, v, hint=None, **style):
    c = WriteOnlyCell(ws, value=v)
    for k, val in style.items():
        setattr(c, k, val)
    if hint:
        c.comment = Comment(hint, AUTHOR, width=320, height=110)
    return c


def write_analysis_report(path, data, progress=lambda p, t: None, mappings=None) -> dict:
    parts, with_role = data.parts, len(data.roles) > 1
    with_total = any(p.has_sums for p in parts)          # сумм нигде нет — нет и «Всего»
    rows = sorted(data.entities, key=lambda e: -e.total)
    wb = Workbook(write_only=True)
    progress(10, "Пишу лист «Одна орг. во всех файлах»…")
    all_inn = write_all_by_inn_sheet(wb, data.paths, mappings)      # первым листом — главный свод
    progress(15, "Пишу лист «Анализ»…")
    ws = wb.create_sheet("Анализ")
    ws.sheet_view.showGridLines = False
    head = ["№", "ИНН", "Название"] + (["Кто"] if with_role else []) + [p.label for p in parts] + \
        ["Всего"] * with_total + ["Есть в файлах", "Файлов", "Ошибки", "Названий"]
    n = len(head)
    err_col = n - 1
    widths = [7, 13, 46] + [12] * with_role + [17] * len(parts) + [18] * with_total + [30, 9, 30, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"D{HEAD_ROW + 1}"
    title_rows(ws, "ГЛУБОКИЙ АНАЛИЗ", "0F172A", f"Создано {datetime.now():%d.%m.%Y %H:%M} · части: "
               + ", ".join(f"«{p.label}»" for p in parts), n)
    _legend(ws)
    hints = _head_hints(parts, with_role)
    ws.append([_cell(ws, t, hints.get(t), font=Font(bold=True), fill=fill_color("E2E8F0"), border=BORDER,
                     alignment=CENTER) for t in head])
    for k, e in enumerate(rows, 1):
        files = present_files(e, parts)
        sums = [(_cell(ws, e.sums[i], number_format=NUM) if p.has_sums else "есть") if e.counts[i] else None
                for i, p in enumerate(parts)]
        ws.append([k, e.inn, _name_cell(ws, e)] + ([e.role] if with_role else []) + sums
                  + [_cell(ws, e.total, number_format=NUM)] * with_total + [", ".join(files), len(files),
                     ", ".join(f for f in FLAGS if f in e.flags),
                     e.names if e.names > 1 else None])
        if k % 20000 == 0:
            progress(15 + int(60 * k / len(rows)), "Пишу лист «Анализ»…")
    last = HEAD_ROW + len(rows)
    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(n)}{last}"
        add_error_color(ws, HEAD_ROW + 1, last, err_col)
        names_col = get_column_letter(n)        # один ИНН — разные названия: название и «Названий» оранжевые
        for rng in (f"C{HEAD_ROW + 1}:C{last}", f"{names_col}{HEAD_ROW + 1}:{names_col}{last}"):
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"${names_col}{HEAD_ROW + 1}>1"],
                                                           fill=cf_fill(ERR_COLOR), font=Font(bold=True)))
    progress(68, "Пишу лист «Сверка со списком»…")
    recon = write_reconcile_sheet(wb, data)
    raw = write_raw_sheets(wb, data.paths, mappings or {})
    progress(70, "Пишу лист «Кто кому платит»…")
    write_pairs_sheet(wb, data)
    write_months_sheet(wb, data)
    progress(80, "Пишу листы с ошибками…")
    write_errors_sheets(wb, data.details)
    write_settings_sheet(wb, data, mappings)
    progress(95, "Сохраняю…")
    wb.save(path)
    return {"path": path, "rows": len(rows), "flags": Counter(f for e in rows for f in e.flags), "reconcile": recon,
            "all_by_inn": all_inn, "raw": raw}


def _legend(ws):
    """Строка 4: цвет ошибок — с подсказкой при наведении."""
    ws.append([None, None, _cell(ws, "Цвет ячейки (наведите — подсказка) →", font=Font(bold=True),
                                 alignment=Alignment(horizontal="right", vertical="center")),
               _cell(ws, "Ошибки", "\n".join(f"• {f}: {FLAG_HINTS[f]}" for f in FLAGS) +
                     "\nПодробно — на листах ошибок (у каждого вида — свой лист).", fill=fill_color(ERR_COLOR),
                     border=BORDER, alignment=CENTER, font=Font(bold=True))])


def _head_hints(parts, with_role) -> dict:
    hints = {"ИНН": "ИНН организации (как записан в файле).",
             "Название": "Самое частое написание названия у этого ИНН.",
             "Кто": "плательщик — кто платил; получатель — кому платили; организация — из обычной таблицы с ИНН.",
             "Всего": "Сумма по всем частям.",
             "Есть в файлах": "В каких файлах есть эта организация (листы одного файла — один файл).",
             "Файлов": "В скольких файлах есть организация. 1 — только в одном файле: фильтром можно отобрать.",
             "Ошибки": "\n".join(f"• {f}: {h}" for f, h in FLAG_HINTS.items()),
             "Названий": "Сколько разных написаний названия у этого ИНН (пусто — одно)."}
    hints.update({p.label: f"Сумма за часть «{p.label}» (файл {p.source})." if p.has_sums else
                  f"В таблице «{p.label}» (файл {p.source}) нет сумм — «есть» значит, что организация там есть."
                  for p in parts})
    return hints
