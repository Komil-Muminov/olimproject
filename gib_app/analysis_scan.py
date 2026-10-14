"""Глубокий анализ: читает любые файлы (реестры платежей и обычные таблицы с ИНН), потом собирает
таблицу с выбранной разбивкой (по файлам / кварталам / месяцам) и парой «было → стало». Чтение — в фоне."""
import os

from analysis_checks import mark_same_names
from analysis_list_map import AS_LIST, as_list, list_guess  # noqa: F401  (AS_LIST — для окна)
from analysis_lists import ListCollector, money_unclear, payments_like
from analysis_model import AnalysisData
from analysis_periods import FILES, available_modes, registry_slots
from analysis_registry import registry_entities, registry_errors, registry_pairs, registry_parts
from excel_com import WriteError
from excel_reader import Book
from registry_read import read_registry
from registry_stats import Registry

ASK_SCORE = 2           # столько узнанных заголовков реестра — спросить, не реестр ли это


def scan_files(paths, mappings=None, progress=lambda p, t: None) -> dict:
    """→ {"raw": (реестры, списки, пути)}, {"needs_inn"/"needs_amount": путь, "guess": …} (ИНН не нашёлся /
    непонятно, где сумма — указать столбцы)
    или {"needs_mapping": путь, "guess": …} (похоже на реестр, но
    столбцы не узнаны — спросить пользователя)."""
    mappings = mappings or {}
    reg, lists, n = Registry(), ListCollector(), len(paths)
    for k, path in enumerate(paths):
        name, base = os.path.basename(path), int(95 * k / n)
        progress(base, f"Читаю {name} ({k + 1} из {n})…")
        sheets, guess = [], None
        if not as_list(mappings.get(path)):
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
        try:
            data = Book(path).sheets
        except Exception as e:
            raise WriteError(f"Не удалось прочитать «{name}»:\n{e}")
        # похоже на реестр (узнаны заголовки или два столбца ИНН), но столбцы не узнаны — спросить
        if guess and path not in mappings and (guess["score"] >= ASK_SCORE or payments_like(data)):
            return {"needs_mapping": path, "guess": guess}
        m = mappings.get(path)
        unclear = money_unclear(data) if m in (None, AS_LIST) else None
        if unclear:                     # несколько столбцов с числами — какой из них сумма?
            return {"needs_amount": path, "guess": unclear}
        found = lists.add_file(path, data, m if isinstance(m, dict) and as_list(m) else None)
        if found:
            continue
        if isinstance(m, dict) and as_list(m):
            raise WriteError(f"В «{name}» в выбранном столбце ИНН нет значений.")
        ask = list_guess(data)          # ИНН не нашёлся сам — пусть пользователь укажет столбец
        if not ask:
            raise WriteError(f"Файл «{name}» пустой.")
        return {"needs_inn": path, "guess": ask}
    progress(97, "Подвожу итоги…")
    if reg.files:
        reg.finish()
    return {"raw": (reg, lists, list(paths))}


def _order(reg, lists, mode) -> tuple:
    slots = registry_slots(reg, mode) if reg.files else []
    parts = registry_parts(slots) + lists.parts
    return slots, parts, sorted(range(len(parts)), key=lambda i: parts[i].order)


def preview_parts(reg, lists, mode: str = FILES) -> list:
    """Части в порядке времени (для выбора «было → стало») — без подсчёта организаций."""
    _, parts, order = _order(reg, lists, mode)
    return [parts[i] for i in order]


def auto_mode(raw) -> str:
    """Разбивка без вопросов: по файлам; если часть получается одна (один файл) — по кварталам/месяцам."""
    reg, lists, _ = raw
    modes = available_modes(reg) if reg.files else [FILES]
    return next((m for m in modes if len(preview_parts(reg, lists, m)) > 1), FILES)


def build(reg, lists, paths, mode: str = FILES, pair=None) -> AnalysisData:
    """Части реестров и списков → одна таблица; части по порядку времени; pair — (было, стало)."""
    slots, parts, order = _order(reg, lists, mode)
    n, details, ents, pairs = len(parts), {}, [], []
    if reg.files:
        details, marked = registry_errors(reg)
        ents = registry_entities(reg, marked, slots)
        pairs = registry_pairs(reg, slots)
        for x in ents + pairs:
            x.sums += [0.0] * len(lists.parts)
            x.counts += [0] * len(lists.parts)
    ents += lists.entities(len(slots), n)
    for x in ents + pairs:
        x.sums = [x.sums[i] for i in order]
        x.counts = [x.counts[i] for i in order]
    details = {**details, **lists.details, "same_name": mark_same_names(ents)}
    compare = pair if pair and n > 1 else ((0, n - 1) if n > 1 else None)
    return AnalysisData([parts[i] for i in order], ents, details, paths, pairs, compare, mode)
