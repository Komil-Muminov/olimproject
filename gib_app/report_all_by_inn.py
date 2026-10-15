"""Лист «Все файлы по ИНН»: строки всех файлов как есть (все колонки), рядом — по ИНН. Только ИНН, которые есть
в двух и более файлах. Каждая колонка подписана: «файл: колонка». Колонку ИНН и строки указал пользователь."""
import os
from collections import Counter

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from analysis_list_map import NAME_KEY, org_key, split_cell
from compare_finder import key_in_text, key_of, match_key
from file_format import real_path
from report_fast import _cell, title_rows
from report_inn_tables import _show, _text
from report_styles import BORDER, CENTER, fill_color

MAX_ROWS = 500_000          # больше — не помещается в память и в разумный Excel: показываются первые
BASE = ["ИНН", "Файл", "Лист", "Строка"]


def _sources(paths: list, mappings: dict) -> list:
    return [(p, m) for p in paths for m in (mappings.get(p) if isinstance(mappings.get(p), list) else [])]


def _rows(path: str, m: dict):
    """(строка Excel, значения) — строки данных листа по настройке пользователя."""
    from python_calamine import CalamineWorkbook

    sheet = CalamineWorkbook.from_path(real_path(path)).get_sheet_by_name(m["sheet"])
    for i, values in enumerate(sheet.iter_rows()):          # строки — с первой строки листа
        if i <= m["head"]:
            continue
        if m.get("last") and i + 1 > m["last"]:
            return
        yield i + 1, values


def _keys(values, m: dict) -> list:
    by_name = bool(m.get("by_name"))
    c = m["cols"].get("name") if by_name else (m["cols"].get("inn") if m["kind"] == "list" else m["cols"].get("payer_inn"))
    out = []
    for v in split_cell(values[c] if c is not None and c < len(values) else None):
        if by_name:
            out += [NAME_KEY + org_key(v)] if org_key(v) else []
        elif (k := key_of(v) or key_in_text(v)):
            out.append(match_key(k) or k)
    return out


def _collect(sources: list) -> tuple:
    """Два прохода: какие ИНН есть в 2+ файлах, затем их строки. → (строки, шапки источников, обрезано ли)."""
    per_file = {}
    for path, m in sources:
        per_file.setdefault(path, set()).update(k for _, values in _rows(path, m) for k in _keys(values, m))
    common = {k for k, n in Counter(k for keys in per_file.values() for k in keys).items() if n >= 2}
    rows, heads = [], []
    for src, (path, m) in enumerate(sources):
        heads.append(_head_row(path, m))
        for r, values in _rows(path, m):
            for k in _keys(values, m):
                if k not in common:
                    continue
                if len(rows) >= MAX_ROWS:
                    return rows, heads + [[]] * (len(sources) - len(heads)), True
                rows.append((k, src, r, [_show(v) if v != "" else None for v in values]))
    return rows, heads, False


def _head_row(path: str, m: dict) -> list:
    head = m.get("labels_row", m["head"])         # строка с названиями колонок (над итогами, если они есть)
    if head < 0:
        return []
    from python_calamine import CalamineWorkbook

    sheet = CalamineWorkbook.from_path(real_path(path)).get_sheet_by_name(m["sheet"])
    for i, values in enumerate(sheet.iter_rows()):
        if i == head:
            return [_text(v) for v in values]
    return []


def write_all_by_inn_sheet(wb, paths: list, mappings: dict) -> dict | None:
    """Только если файлов два и больше и столбцы указаны пользователем. → {"orgs", "rows", "cut"} или None."""
    sources = _sources(paths, mappings or {})
    if len({p for p, _ in sources}) < 2:
        return None
    rows, heads, cut = _collect(sources)
    widths = [len(heads[s]) for s in range(len(sources))]
    for _, src, _, values in rows:
        widths[src] = max(widths[src], len(values))
    ws = wb.create_sheet("Все файлы по ИНН")
    ws.sheet_properties.tabColor = "7C3AED"
    ws.freeze_panes = "B5"
    labels = BASE + [f"{os.path.basename(p)}: {get_column_letter(m['col0'] + i + 1)}"
                     + (f" {heads[s][i]}" if i < len(heads[s]) and heads[s][i] else "")
                     for s, (p, m) in enumerate(sources) for i in range(widths[s])]
    note = f" Показаны первые {MAX_ROWS:,} строк — сузьте строки в поле ИНН.".replace(",", " ") if cut else ""
    title_rows(ws, f"ВСЕ ФАЙЛЫ ПО ИНН — {len({k for k, *_ in rows}):,}".replace(",", " "), "7C3AED",
               "Организации, которые есть в двух и более файлах: все их строки как есть, рядом. Колонки подписаны "
               "«файл: колонка»." + note, len(labels))
    ws.append([_cell(ws, t, font=Font(bold=True), fill=fill_color("E2E8F0"), border=BORDER, alignment=CENTER) for t in labels])
    offsets = [len(BASE) + sum(widths[:s]) for s in range(len(sources))]
    for k, src, r, values in sorted(rows, key=lambda x: (x[0], x[1], x[2])):
        path, m = sources[src]
        line = [None] * len(labels)
        line[:4] = [k[1:] if k.startswith(NAME_KEY) else (k.zfill(9) if k.isdigit() and len(k) < 9 else k),
                    os.path.basename(path), m["sheet"], r]
        line[offsets[src]:offsets[src] + len(values)] = values
        ws.append(line)
    ws.auto_filter.ref = f"A4:{get_column_letter(len(labels))}{4 + len(rows)}"
    return {"orgs": len({k for k, *_ in rows}), "rows": len(rows), "cut": cut}
