"""«Один ИНН во всех файлах»: все строки, где встречается этот ИНН — в любой колонке, в том числе внутри текста
(«… ИНН 040058284»), из всех взятых файлов (лист и строки — из карточек); можно с порогом суммы.
Отдельный Excel: где нашлось и сумма по файлам, ниже строки как есть."""
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


def write_inn_report(path: str, paths: list, mappings: dict, inn: str, threshold=None,
                     progress=lambda p, t: None) -> dict:
    """→ {"path" (None — не найден), "inn", "rows", "files": [(таблица, строк, сумма)]}; threshold — (от, до)."""
    pattern = inn_pattern(inn)
    found = collect(paths, mappings, lambda values, m: has_inn(values, pattern) and in_range(row_sum(values, m), threshold),
                    progress)
    res = {"path": None, "inn": inn, "rows": sum(x[1] for x in found["per"]),
           "files": [x for x in found["per"] if x[1]], "threshold": threshold}
    if not res["rows"]:
        return res
    progress(92, "Сохраняю…")
    wb = Workbook(write_only=True)
    limit = f", суммы {threshold_text(threshold)}" if threshold else ""
    write_rows_sheet(wb, f"ИНН {inn}", f"ИНН {inn}{limit} — {res['rows']:,} строк в {len(res['files'])} файлах".replace(",", " "),
                     "Все строки этого ИНН из выбранных файлов, как есть. Над каждой колонкой — из какого она файла.", found)
    wb.save(path)
    res["path"] = path
    return res
