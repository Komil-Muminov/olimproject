"""«Один ИНН во всех файлах»: все строки этого ИНН из всех взятых файлов (лист, колонки и строки — из карточек),
без предела по числу строк. Сверху — где нашлось и сумма по колонкам «Сумма»; ниже — строки как есть,
каждая колонка подписана «файл: колонка»."""
import os

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from compare_finder import match_key
from report_all_by_inn import _head_row, _keys, _rows, _show, _sources, _table_name
from report_fast import _cell, title_rows
from report_styles import BORDER, CENTER, NUM, fill_color
from utils import to_num_safe

BASE = ["Файл", "Лист", "Строка"]


def clean_inn(text: str) -> str:
    """ИНН из введённого: только цифры; потерянный ноль впереди возвращается (40058284 → 040058284)."""
    digits = "".join(ch for ch in str(text or "") if ch.isdigit())
    return digits.zfill(9) if 0 < len(digits) < 9 else digits


def _head(ws, texts: list):
    ws.append([_cell(ws, t, font=Font(bold=True), fill=fill_color("E2E8F0"), border=BORDER, alignment=CENTER) for t in texts])


def write_inn_report(path: str, paths: list, mappings: dict, inn: str, progress=lambda p, t: None) -> dict:
    """→ {"path" (None — не найден), "inn", "rows", "files": [(таблица, строк, сумма)]}."""
    target, sources = match_key(inn), _sources(paths, mappings or {})
    hits, heads, per = [], [], []
    for src, (p, m) in enumerate(sources):
        progress(int(90 * src / max(len(sources), 1)), f"Ищу в {os.path.basename(p)}…")
        heads.append(_head_row(p, m))
        found = [(r, values) for r, values in _rows(p, m) if target in _keys(values, m)]
        money = m.get("sums") or []
        per.append((_table_name(p, m, sources), len(found),
                    round(sum(to_num_safe(v[c]) for _, v in found for c in money if c < len(v)), 2) if money else None))
        hits += [(src, r, [_show(v) if v != "" else None for v in values]) for r, values in found]
    res = {"path": None, "inn": inn, "rows": len(hits), "files": [x for x in per if x[1]]}
    if not hits:
        return res
    progress(92, "Сохраняю…")
    widths = [len(h) for h in heads]
    for src, _, values in hits:
        widths[src] = max(widths[src], len(values))
    labels = BASE + [f"{per[s][0]}: {get_column_letter(m['col0'] + i + 1)}" + (f" {heads[s][i]}" if i < len(heads[s]) and heads[s][i] else "")
                     for s, (_, m) in enumerate(sources) for i in range(widths[s])]
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(f"ИНН {inn}")
    ws.sheet_properties.tabColor = "7C3AED"
    title_rows(ws, f"ИНН {inn} — {len(hits):,} строк в {len(res['files'])} файлах".replace(",", " "), "7C3AED",
               "Все строки этого ИНН из выбранных файлов, как есть. Над каждой колонкой — из какого она файла.", len(labels))
    _head(ws, ["Где нашлось", "Строк", "Сумма (колонки «Сумма»)"])
    for name, count, total in per:
        cell = WriteOnlyCell(ws, value=total)
        cell.number_format = NUM
        ws.append([name, count, cell if total is not None else "не указана"])
    ws.append([])
    _head(ws, labels)
    offsets = [len(BASE) + sum(widths[:s]) for s in range(len(sources))]
    for src, r, values in hits:
        p, m = sources[src]
        line = [None] * len(labels)
        line[:len(BASE)] = [os.path.basename(p), m["sheet"], r]
        line[offsets[src]:offsets[src] + len(values)] = values
        ws.append(line)
    wb.save(path)
    res["path"] = path
    return res
