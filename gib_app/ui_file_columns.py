"""Для окна «Как читать файл»: подписи столбцов с примерами и проверка настроек листа."""
from openpyxl.utils import get_column_letter

from analysis_list_map import AS_LIST
from analysis_propose import REGISTRY
from constants import SUM_ROW_LABELS
from utils import format_money, norm_text

FIELDS = {REGISTRY: [("payer_inn", "ИНН плательщика"), ("payer", "Плательщик"), ("rec_inn", "ИНН получателя"),
                     ("rec", "Получатель"), ("date", "Дата"), ("doc", "№ документа"), ("purpose", "Основание")],
          AS_LIST: [("inn", "ИНН  *"), ("name", "Название")]}


def text(r, i) -> str:
    """Ячейка для подписи: число — «1 234,56», текст — первые 30 знаков."""
    v = r[i] if i < len(r) else ""
    if isinstance(v, float):
        return format_money(v)
    return str(v).strip()[:30]


def sample_row(rows, h: int) -> list:
    """Пример значений — первая строка под заголовками, кроме строк итогов («Ҳамагӣ», «Итого»)."""
    for r in rows[h + 1:]:
        if not any(isinstance(v, str) and norm_text(v).startswith(SUM_ROW_LABELS) for v in r):
            return r
    return []


def column_names(sheet: dict, h: int) -> list:
    """«C: РМА   (напр. 040058284)» — по строке заголовков h (с 0) листа из propose()."""
    rows = sheet["rows"]
    heads, sample = rows[h], sample_row(rows, h)
    return [f"{get_column_letter(sheet['col0'] + i + 1)}: {text(heads, i) or '(без названия)'}"
            + (f"   (напр. {text(sample, i)})" if text(sample, i) else "") for i in range(max(len(r) for r in rows))]


def problem(st: dict) -> str:
    """Чего не хватает в настройке листа ("" — всё есть)."""
    cols = st["cols"]
    if st["kind"] == REGISTRY:
        if cols.get("payer_inn") is None and cols.get("rec_inn") is None:
            return "нужен хотя бы один ИНН (плательщика или получателя)"
        return "" if st["sums"] else "нужен хотя бы один столбец суммы"
    return "" if cols.get("inn") is not None else "укажите столбец с ИНН"


def to_mapping(name: str, sheet: dict, st: dict) -> dict:
    """Настройка листа → для scan_files: реестр — столбцы + "amount" (список); список — "kind": "list", "sums"."""
    m = {"sheet": name, "head": st["head"], "row0": sheet["row0"], "col0": sheet["col0"], "kind": st["kind"],
         "cols": dict(st["cols"])}
    if st["kind"] == REGISTRY:
        m["cols"]["amount"] = list(st["sums"])
    else:
        m["sums"] = list(st["sums"])
    return m
