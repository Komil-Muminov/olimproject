"""Фильтр по всем взятым файлам: ИНН (в любой колонке строки, в том числе внутри текста) и/или порог суммы
строки (по колонкам «Сумма» из карточки). Отдельный Excel: где нашлось и сумма по файлам, ниже строки как есть."""
import re

from openpyxl import Workbook

from compare_finder import match_key
from report_all_by_inn import _show
from report_rows import collect, in_range, row_sum, threshold_text, write_rows_sheet


def clean_inn(text: str) -> str:
    """ИНН из введённого: только цифры; потерянный ноль впереди возвращается (40058284 → 040058284)."""
    digits = "".join(ch for ch in str(text or "") if ch.isdigit())
    return digits.zfill(9) if 0 < len(digits) < 9 else digits


def inn_pattern(inn: str):
    """ИНН отдельным числом или внутри текста; нули впереди могли потеряться (040058284 = 40058284)."""
    return re.compile(rf"(?<!\d)0*{re.escape(match_key(inn))}(?!\d)")


def has_inn(values, pattern) -> bool:
    """ИНН в любой ячейке строки."""
    return any(pattern.search(str(_show(v))) for v in values if v not in (None, ""))


def filter_title(inn: str = "", threshold=None) -> str:
    """«ИНН 040058284», «Суммы от 2 000,00» или вместе."""
    return ", ".join(x for x in (f"ИНН {inn}" if inn else "", f"суммы {threshold_text(threshold)}" if threshold else "") if x)


def write_filter_report(path: str, paths: list, mappings: dict, inn: str = "", threshold=None,
                        progress=lambda p, t: None) -> dict:
    """→ {"path" (None — ничего не нашлось), "title", "rows", "files": [(таблица, строк, сумма)]}."""
    pattern = inn_pattern(inn) if inn else None

    def keep(values, m) -> bool:
        if pattern and not has_inn(values, pattern):
            return False
        return in_range(row_sum(values, m), threshold)

    found = collect(paths, mappings, keep, progress)
    title = filter_title(inn, threshold)
    res = {"path": None, "title": title, "rows": sum(x[1] for x in found["per"]),
           "files": [x for x in found["per"] if x[1]]}
    if not res["rows"]:
        return res
    progress(92, "Сохраняю…")
    wb = Workbook(write_only=True)
    write_rows_sheet(wb, title[:31], f"{title.upper()} — {res['rows']:,} строк в {len(res['files'])} файлах".replace(",", " "),
                     "Строки из выбранных файлов, как есть: ИНН — в любой колонке, сумма строки — по колонкам «Сумма». "
                     "Над каждой колонкой — из какого она файла.", found)
    wb.save(path)
    res["path"] = path
    return res
