"""Фильтр по всем взятым файлам: ИНН через запятую (в любой колонке строки, в том числе внутри текста) и/или порог
суммы строки (по колонкам «Сумма» из карточки). Отдельный Excel: где нашлось и сумма — по файлам и по каждому ИНН,
ниже строки как есть."""
import re

from openpyxl import Workbook

from inn_keys import match_key
from report_rows import collect, in_range, row_sum, threshold_text, write_rows_sheet
from report_source import show

TITLE_INNS = 3          # больше ИНН — в названии файла только их число


def clean_inn(text: str) -> str:
    """ИНН из введённого: только цифры; потерянный ноль впереди возвращается (40058284 → 040058284)."""
    digits = "".join(ch for ch in str(text or "") if ch.isdigit())
    return digits.zfill(9) if 0 < len(digits) < 9 else digits


def parse_inns(text: str) -> tuple:
    """«040058284, 630000962» → ([ИНН без повторов], [неправильные куски])."""
    good, bad = [], []
    for part in re.split(r"[,;\n]+", text or ""):
        digits = sum(ch.isdigit() for ch in part)
        if part.strip():
            good.append(clean_inn(part)) if 8 <= digits <= 14 else bad.append(part.strip())
    return list(dict.fromkeys(good)), bad


def inn_pattern(inns: list):
    """Любой из ИНН отдельным числом или внутри текста; нули впереди могли потеряться (040058284 = 40058284)."""
    keys = "|".join(sorted({re.escape(match_key(i)) for i in inns}, key=len, reverse=True))
    return re.compile(rf"(?<!\d)0*({keys})(?!\d)")


def inns_in_row(values, pattern, by_key: dict) -> list:
    """Какие из искомых ИНН есть в строке (в любой ячейке)."""
    found = {m.group(1) for v in values if v not in (None, "") for m in pattern.finditer(str(show(v)))}
    return [inn for key, inn in by_key.items() if key in found]


def filter_title(inns: list = (), threshold=None) -> str:
    """«ИНН 040058284», «ИНН 1, 2», «5 ИНН», «суммы от 2 000,00» или вместе."""
    inn = (f"ИНН {', '.join(inns)}" if len(inns) <= TITLE_INNS else f"{len(inns)} ИНН") if inns else ""
    return ", ".join(x for x in (inn, f"суммы {threshold_text(threshold)}" if threshold else "") if x)


def write_filter_report(path: str, paths: list, mappings: dict, inns: list = (), threshold=None,
                        progress=lambda p, t: None) -> dict:
    """→ {"path" (None — ничего не нашлось), "title", "rows", "files": [(таблица, строк, сумма)],
    "missing": [ИНН, которых нет ни в одном файле]}."""
    by_key = {match_key(i): i for i in inns}
    pattern = inn_pattern(inns) if inns else None

    def keep(values, m):
        if not in_range(row_sum(values, m), threshold):
            return False
        return inns_in_row(values, pattern, by_key) if pattern else True

    found = collect(paths, mappings, keep, progress)
    title = filter_title(inns, threshold)
    seen = {tag for tag, _ in found["by_tag"]}
    res = {"path": None, "title": title, "rows": sum(x[1] for x in found["per"]),
           "files": [x for x in found["per"] if x[1]], "missing": [i for i in inns if i not in seen]}
    if not res["rows"]:
        return res
    progress(92, "Сохраняю…")
    wb = Workbook(write_only=True)
    write_rows_sheet(wb, title[:31], f"{title.upper()} — {res['rows']:,} строк в {len(res['files'])} файлах".replace(",", " "),
                     "Строки из выбранных файлов, как есть: ИНН — в любой колонке, сумма строки — по колонкам «Сумма». "
                     "Над каждой колонкой — из какого она файла.", found, list(inns) if len(inns) > 1 else None)
    wb.save(path)
    res["path"] = path
    return res
