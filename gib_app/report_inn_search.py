"""«Один ИНН во всех файлах»: все строки этого ИНН из всех взятых файлов (лист, колонки и строки — из карточек),
можно с порогом суммы. Отдельный Excel: где нашлось и сумма по файлам, ниже строки как есть."""
from openpyxl import Workbook

from compare_finder import match_key
from report_all_by_inn import _keys
from report_rows import collect, in_range, row_sum, threshold_text, write_rows_sheet


def clean_inn(text: str) -> str:
    """ИНН из введённого: только цифры; потерянный ноль впереди возвращается (40058284 → 040058284)."""
    digits = "".join(ch for ch in str(text or "") if ch.isdigit())
    return digits.zfill(9) if 0 < len(digits) < 9 else digits


def write_inn_report(path: str, paths: list, mappings: dict, inn: str, threshold=None,
                     progress=lambda p, t: None) -> dict:
    """→ {"path" (None — не найден), "inn", "rows", "files": [(таблица, строк, сумма)]}; threshold — (от, до)."""
    target = match_key(inn)
    found = collect(paths, mappings, lambda values, m: target in _keys(values, m) and in_range(row_sum(values, m), threshold),
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
