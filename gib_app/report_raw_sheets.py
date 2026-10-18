"""Листы «без ИНН» — в свод как есть: значения листа без изменений, колонки на тех же местах (A, B, C…).
Имя листа в своде — «файл · лист» (Excel разрешает до 31 знака и без []:*?/\\)."""
import os
import re

from analysis_list_map import RAW
from file_format import real_path

MAX_ROWS = 500_000
BAD_CHARS = re.compile(r"[\[\]:*?/\\]")


def _title(path: str, sheet: str, taken: set) -> str:
    """Уникальное имя листа: «файл · лист», обрезанное до 31 знака; занято — с номером."""
    base = BAD_CHARS.sub("_", f"{os.path.splitext(os.path.basename(path))[0]} · {sheet}")[:31]
    name, k = base, 2
    while name.lower() in taken:
        suffix = f" ({k})"
        name, k = base[:31 - len(suffix)] + suffix, k + 1
    taken.add(name.lower())
    return name


def write_raw_sheets(wb, paths: list, mappings: dict) -> list:
    """→ [(файл, лист, имя листа в своде, строк)] — что скопировано как есть."""
    from python_calamine import CalamineWorkbook

    out, taken = [], {s.lower() for s in wb.sheetnames}
    for path in paths:
        chosen = mappings.get(path) if isinstance(mappings.get(path), list) else []
        raw = [m["sheet"] for m in chosen if m["kind"] == RAW]
        if not raw:
            continue
        book = CalamineWorkbook.from_path(real_path(path))
        for sheet_name in raw:
            sheet = book.get_sheet_by_name(sheet_name)
            pad = [None] * (sheet.start or (0, 0))[1]       # iter_rows — с первой заполненной колонки
            ws = wb.create_sheet(_title(path, sheet_name, taken))
            ws.sheet_properties.tabColor = "94A3B8"
            n = 0
            for n, values in enumerate(sheet.iter_rows(), 1):
                if n > MAX_ROWS:
                    break
                ws.append(pad + [None if v == "" else v for v in values])
            out.append((os.path.basename(path), sheet_name, ws.title, min(n, MAX_ROWS)))
    return out
