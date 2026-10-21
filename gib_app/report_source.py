"""Строки таблиц по настройке из карточек: таблица — лист файла (несколько листов одного файла — разные таблицы).
Значения — как есть: дата текстом, целое число без «.0»."""
import os
from datetime import date, datetime

from file_format import real_path


def text(v) -> str:
    return " ".join(str(v).split()) if v not in (None, "") else ""


def show(v):
    """Значение для отчёта: дата — текстом, целое число — без «.0»."""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def sources(paths: list, mappings: dict) -> list:
    """[(файл, настройка листа)] — по каждой настройке из карточек."""
    return [(p, m) for p in paths for m in (mappings.get(p) if isinstance(mappings.get(p), list) else [])]


def _sheet(path: str, name: str):
    from python_calamine import CalamineWorkbook

    return CalamineWorkbook.from_path(real_path(path)).get_sheet_by_name(name)


def data_rows(path: str, m: dict):
    """(строка Excel, значения) — строки данных листа: ниже m["head"], до m["last"] (если указан)."""
    for i, values in enumerate(_sheet(path, m["sheet"]).iter_rows()):          # строки — с первой строки листа
        if i <= m["head"]:
            continue
        if m.get("last") and i + 1 > m["last"]:
            return
        yield i + 1, values


def head_row(path: str, m: dict) -> list:
    """Названия колонок (строка над данными, над итогами — если они есть)."""
    head = m.get("labels_row", m["head"])
    if head < 0:
        return []
    for i, values in enumerate(_sheet(path, m["sheet"]).iter_rows()):
        if i == head:
            return [text(v) for v in values]
    return []


def table_name(path: str, m: dict, all_sources: list) -> str:
    """Подпись таблицы: имя файла; у файла несколько листов — ещё и лист."""
    name = os.path.basename(path)
    return f"{name} · {m['sheet']}" if sum(p == path for p, _ in all_sources) > 1 else name
