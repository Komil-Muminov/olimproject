"""Для окна «Как читать файл»: подписи столбцов с примерами и проверка настроек листа."""
from openpyxl.utils import get_column_letter

from analysis_list_map import AS_LIST, transpose
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


def sheet_view(sheet: dict, flip: bool) -> dict:
    """Лист как видит окно; flip — «перевёрнутая» таблица: строки ↔ столбцы (строки — столбцы листа)."""
    if not flip:
        return sheet
    return dict(sheet, rows=transpose(sheet["rows"]), row0=sheet["col0"], col0=0, flip=True)


def column_names(sheet: dict, h: int) -> list:
    """«C: РМА   (напр. 040058284)» — по строке заголовков h (с 0) листа из propose()."""
    rows = sheet["rows"]
    heads, sample = rows[h], sample_row(rows, h)
    where = (lambda i: f"стр. {i + 1}") if sheet.get("flip") else (lambda i: get_column_letter(sheet["col0"] + i + 1))
    return [f"{where(i)}: {text(heads, i) or '(без названия)'}"
            + (f"   (напр. {text(sample, i)})" if text(sample, i) else "") for i in range(max(len(r) for r in rows))]


def problem(st: dict) -> str:
    """Чего не хватает в настройке листа ("" — всё есть)."""
    cols = st["cols"]
    if st["kind"] == REGISTRY:
        if cols.get("payer_inn") is None and cols.get("rec_inn") is None:
            return "нужен хотя бы один ИНН (плательщика или получателя)"
        return "" if st["sums"] else "нужен хотя бы один столбец суммы"
    if st.get("by_name"):
        return "" if cols.get("name") is not None else "укажите столбец с названием (ИНН нет)"
    return "" if cols.get("inn") is not None else "укажите столбец с ИНН"


def formula_warning(name: str, sheet: dict, st: dict) -> str:
    """Отмеченные столбцы сумм с формулами без сохранённого результата — суммы выйдут пустыми."""
    bad = [f"{get_column_letter(sheet['col0'] + c + 1)} ({n})" for c, n in sheet.get("formulas", {}).items()
           if c in st["sums"] and not st.get("transpose")]
    return (f"«{name}»: в столбцах {', '.join(bad)} формулы без сохранённого результата — суммы будут пустыми. "
            "Откройте файл в Excel и сохраните его.") if bad else ""


def to_mapping(name: str, sheet: dict, st: dict) -> dict:
    """Настройка листа → для scan_files: реестр — столбцы + "amount" (список); список — "kind": "list", "sums".
    sheet — лист так, как его видит окно (sheet_view): для «перевёрнутой» таблицы сдвиги уже свои."""
    m = {"sheet": name, "head": st["head"], "row0": sheet["row0"], "col0": sheet["col0"], "kind": st["kind"],
         "cols": dict(st["cols"]), "transpose": bool(st.get("transpose")), "by_name": bool(st.get("by_name")),
         "separate": bool(st.get("separate")), "repeats": bool(st.get("repeats")),
         "hidden_sheet": bool(sheet.get("hidden"))}
    names = [n.split("   (напр.")[0] for n in column_names(sheet, st["head"])]     # «C: РМА» — для «Как считали»
    m["labels"] = {title.rstrip(" *"): names[st["cols"][key]] for key, title in FIELDS[st["kind"]]
                   if st["cols"].get(key) is not None and st["cols"][key] < len(names)}
    m["sum_labels"] = [names[c] for c in st["sums"] if c < len(names)]
    if st["kind"] == REGISTRY:
        m["cols"]["amount"] = list(st["sums"])
    else:
        m["sums"] = list(st["sums"])
    return m
