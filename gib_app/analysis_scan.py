"""Глубокий анализ: читает листы с ИНН (как настроил пользователь в окне «Листы и ИНН») и собирает одну
таблицу: организация — строка, каждый лист — часть. Чтение — в фоне."""
import os

from analysis_checks import inn_names, mark_same_names
from analysis_list_map import RAW
from analysis_lists import ListCollector
from analysis_model import AnalysisData
from excel_com import WriteError
from excel_reader import Book
from excel_sheet_info import total_check, white_cells


def _read_chosen(path, sheet_maps: list, lists):
    """Листы, настроенные пользователем. Листы «без ИНН» не читаются — они попадут в свод как есть
    (report_raw_sheets)."""
    sheet_maps = [m for m in sheet_maps if m["kind"] != RAW]
    if not sheet_maps:
        return
    data, found = Book(path).sheets, 0
    for m in sheet_maps:
        found += lists.add_file(path, data, m)
        check = total_check(path, m, lists.last_total)
        lists.checks += [check] if check else []
        lists.white_cells += [[os.path.basename(path), m["sheet"], ref, text] for ref, text in white_cells(path, m["sheet"])]
    if not found:
        raise WriteError(f"В «{os.path.basename(path)}» в выбранных столбцах ИНН нет значений.")


def scan_files(paths, mappings=None, progress=lambda p, t: None) -> dict:
    """mappings[путь] — список настроек листов (окно «Листы и ИНН»). Без настроек — таблицы с ИНН ищутся
    сами. → {"raw": (списки, пути)}."""
    mappings, lists, n = mappings or {}, ListCollector(), len(paths)
    for k, path in enumerate(paths):
        name = os.path.basename(path)
        progress(int(95 * k / n), f"Читаю {name} ({k + 1} из {n})…")
        if isinstance(mappings.get(path), list):
            _read_chosen(path, mappings[path], lists)
            continue
        try:
            data = Book(path).sheets
        except Exception as e:
            raise WriteError(f"Не удалось прочитать «{name}»:\n{e}")
        if not lists.add_file(path, data):
            raise WriteError(f"В «{name}» не найдено таблиц с ИНН.")
    progress(97, "Подвожу итоги…")
    return {"raw": (lists, list(paths))}


def build(lists, paths) -> AnalysisData:
    """Части (листы) → одна таблица; части по порядку времени."""
    parts = lists.parts
    order = sorted(range(len(parts)), key=lambda i: parts[i].order)
    n, ents = len(parts), lists.entities(0, len(parts))
    for e in ents:
        e.sums = [e.sums[i] for i in order]
        e.counts = [e.counts[i] for i in order]
        e.cols = [e.cols[i] for i in order] if e.cols else [None] * n
    details = {**lists.details, "same_name": mark_same_names(ents), "inn_names": inn_names(ents),
               "white_text": lists.white_cells}
    return AnalysisData([parts[i] for i in order], ents, details, paths, lists.notes, lists.checks)
