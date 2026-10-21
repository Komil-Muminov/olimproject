"""Глубокий анализ: читает любые файлы (реестры платежей и обычные таблицы с ИНН), потом собирает
таблицу с выбранной разбивкой (по файлам / кварталам / месяцам) и парой «было → стало». Чтение — в фоне."""
import os

from analysis_checks import inn_names, mark_same_names
from analysis_list_map import AS_LIST, as_list, list_guess  # noqa: F401  (AS_LIST — для окна)
from analysis_lists import ListCollector, money_unclear, payments_like
from analysis_model import AnalysisData
from analysis_periods import FILES, available_modes, registry_slots
from analysis_registry import registry_entities, registry_errors, registry_pairs, registry_parts
from excel_com import WriteError
from excel_reader import Book, sheet_values
from excel_sheet_info import hidden_note, sheet_scan, white_cells
from registry_read import read_registry
from registry_stats import Registry

ASK_SCORE = 2           # столько узнанных заголовков реестра — спросить, не реестр ли это


def _add_registry(reg, path, sheets, tick, hidden=frozenset()):
    """hidden — скрытые строки листа (если их немного): такие платежи попадут на лист «Скрытые строки»."""
    for sheet, total, rows in sheets:
        f = reg.start_file(path, sheet)
        for i, p in enumerate(rows):
            reg.add(f, p)
            if hidden and p.row in hidden:
                reg.hidden.append((f, p))
            if i % 50000 == 0:
                tick(i, total)


def _read_chosen(path, sheet_maps: list, reg, lists, tick):
    """Листы, настроенные пользователем (окно «Как читать файл»): реестры и списки, каждый лист — по-своему."""
    found = 0
    for m in sheet_maps:
        if m["kind"] == AS_LIST:            # один нужный лист — быстро (реестр на 800 тыс. строк тоже)
            found += lists.add_file(path, {m["sheet"]: sheet_values(path, m["sheet"])}, m)
        else:
            hidden = sheet_scan(path, m["sheet"])["hidden_rows"]
            note = hidden_note(path, m["sheet"], hidden)
            reg.notes += [note] if note else []
            sheets, _ = read_registry(path, m)
            _add_registry(reg, path, sheets, tick, hidden if note and note["listed"] else frozenset())
            found += len(sheets)
        reg.white_cells += [[os.path.basename(path), m["sheet"], ref, text] for ref, text in white_cells(path, m["sheet"])]
    if not found:
        raise WriteError(f"В «{os.path.basename(path)}» в выбранных столбцах ИНН нет значений.")


def scan_files(paths, mappings=None, progress=lambda p, t: None) -> dict:
    """mappings[путь] — список настроек листов (окно «Как читать файл») или прежние одиночные настройки.
    → {"raw": (реестры, списки, пути)}; без настроек файл узнаётся сам и может вернуться
    {"needs_inn" / "needs_amount" / "needs_mapping": путь, "guess": …} — спросить пользователя."""
    mappings = mappings or {}
    reg, lists, n = Registry(), ListCollector(), len(paths)
    for k, path in enumerate(paths):
        name, base = os.path.basename(path), int(95 * k / n)
        progress(base, f"Читаю {name} ({k + 1} из {n})…")
        tick = lambda i, total: progress(base + int(95 / n * i / max(total, 1)), f"Считаю {name} ({k + 1} из {n})…")
        if isinstance(mappings.get(path), list):
            _read_chosen(path, mappings[path], reg, lists, tick)
            continue
        sheets, guess = [], None
        if not as_list(mappings.get(path)):
            try:
                sheets, guess = read_registry(path, mappings.get(path))
            except Exception:
                pass
        if sheets:
            _add_registry(reg, path, sheets, tick)
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
    for e in ents:                      # значения столбцов списка — только у организаций из списков
        e.cols = [e.cols[i] for i in order] if e.cols else [None] * n
    details = {**details, **lists.details, "same_name": mark_same_names(ents), "inn_names": inn_names(ents),
               "white_text": reg.white_cells}
    compare = pair if pair and n > 1 else ((0, n - 1) if n > 1 else None)
    if reg.files and lists.parts:       # реестр + список — разные вещи: сравнения нет (для них — «Сверка»)
        compare = None
    if len({p.source for p in parts}) > 2 and not pair:     # больше двух файлов — суммы рядом, без «было → стало»
        compare = None
    months = dict(sorted(reg.months.items())) if reg.files else {}
    return AnalysisData([parts[i] for i in order], ents, details, paths, pairs, compare, mode, months,
                        reg.notes + lists.notes)
