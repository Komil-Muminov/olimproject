"""Глубокий анализ: читает любые файлы (реестры платежей и обычные таблицы с ИНН) в одну таблицу.
Запускается в фоновом потоке."""
import os

from analysis_lists import ListCollector
from analysis_model import AnalysisData
from analysis_registry import registry_entities, registry_errors, registry_parts
from excel_com import WriteError
from excel_reader import Book
from registry_read import read_registry
from registry_stats import Registry

AS_LIST = "list"        # пользователь сказал: читать файл как обычный список организаций
ASK_SCORE = 2           # столько узнанных заголовков реестра — спросить, не реестр ли это


def scan_files(paths, mappings=None, progress=lambda p, t: None) -> dict:
    """→ {"data": AnalysisData} или {"needs_mapping": путь, "guess": …} (похоже на реестр, но столбцы
    не узнаны — спросить пользователя)."""
    mappings = mappings or {}
    reg, lists, n = Registry(), ListCollector(), len(paths)
    for k, path in enumerate(paths):
        name, base = os.path.basename(path), int(95 * k / n)
        progress(base, f"Читаю {name} ({k + 1} из {n})…")
        sheets, guess = [], None
        if mappings.get(path) != AS_LIST:
            try:
                sheets, guess = read_registry(path, mappings.get(path))
            except Exception:
                pass
        if sheets:
            for sheet, total, rows in sheets:
                f = reg.start_file(path, sheet)
                for i, p in enumerate(rows):
                    reg.add(f, p)
                    if i % 50000 == 0:
                        progress(base + int(95 / n * i / max(total, 1)), f"Считаю {name} ({k + 1} из {n})…")
            continue
        if guess and guess["score"] >= ASK_SCORE and path not in mappings:
            return {"needs_mapping": path, "guess": guess}
        try:
            found = lists.add_file(path, Book(path).sheets)
        except Exception as e:
            raise WriteError(f"Не удалось прочитать «{name}»:\n{e}")
        if not found:
            raise WriteError(f"В «{name}» не нашлось ни платежей, ни таблиц с ИНН.")
    progress(97, "Подвожу итоги…")
    if reg.files:
        reg.finish()
    return {"data": build(reg, lists, list(paths))}


def build(reg, lists, paths) -> AnalysisData:
    """Части реестров и списков → одна таблица; части по порядку времени."""
    rparts = registry_parts(reg) if reg.files else []
    parts = rparts + lists.parts
    n, details, ents = len(parts), [], []
    if reg.files:
        details, marked = registry_errors(reg)
        ents = registry_entities(reg, marked)
        for e in ents:
            e.sums += [0.0] * len(lists.parts)
            e.counts += [0] * len(lists.parts)
    ents += lists.entities(len(rparts), n)
    order = sorted(range(n), key=lambda i: parts[i].order)
    for e in ents:
        e.sums = [e.sums[i] for i in order]
        e.counts = [e.counts[i] for i in order]
    return AnalysisData([parts[i] for i in order], ents, details + lists.details, paths)
