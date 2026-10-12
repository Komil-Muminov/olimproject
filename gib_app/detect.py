"""Автопоиск на листе: столбец с назначениями, строка шапки ГИБ, столбец суммы."""
from dataclasses import dataclass, field

from excel_io import Book, gib_code
from parser import parse_purpose, to_number


@dataclass
class Layout:
    purpose_col: int
    first_row: int
    last_row: int
    header_row: int | None          # None — шапки нет, её нужно вставить над данными
    amount_col: int | None = None
    gib_cols: dict = field(default_factory=dict)   # код ГИБ -> столбец (уже есть в шапке)


def best_purpose_column(data):
    """Столбец, где больше всего назначений с ГИБ. Возвращает (столбец, первая, последняя строка)."""
    counts = {}
    for r, row in enumerate(data, 1):
        for c, v in enumerate(row, 1):
            if isinstance(v, str) and parse_purpose(v):
                n, first, _ = counts.get(c, (0, r, r))
                counts[c] = (n + 1, first, r)
    if not counts:
        return None
    col = max(counts, key=lambda c: counts[c][0])
    _, first, last = counts[col]
    return col, first, last


def range_from(data, col, start_row):
    """Назначения от выбранной ячейки вниз — до последней строки с ГИБ в этом столбце."""
    last = start_row
    for r in range(start_row, len(data) + 1):
        v = Book.value(data, r, col)
        if isinstance(v, str) and parse_purpose(v):
            last = r
    return last


def header_row_above(data, col, first_row):
    """Шапка ГИБ — ближайшая строка выше данных с кодами ГИБ; иначе ближайшая непустая строка."""
    if first_row <= 1:
        return None
    for r in range(first_row - 1, max(0, first_row - 8), -1):
        if any(gib_code(v) for v in data[r - 1]):
            return r
    for r in range(first_row - 1, max(0, first_row - 8), -1):
        if any(v not in (None, "") for v in data[r - 1]):
            return r
    return first_row - 1


def amount_column(data, col, first_row, last_row):
    """Столбец суммы платежа: где число совпадает с суммой ГИБ в большинстве строк."""
    hits = {}
    total = 0
    for r in range(first_row, last_row + 1):
        gibs = parse_purpose(Book.value(data, r, col) or "")
        if not gibs:
            continue
        total += 1
        s = round(sum(gibs.values()), 2)
        row = data[r - 1]
        for c, v in enumerate(row, 1):
            n = to_number(v)
            if c != col and n is not None and abs(n - s) < 0.005:
                hits[c] = hits.get(c, 0) + 1
    if not hits or total == 0:
        return None
    c = max(hits, key=hits.get)
    return c if hits[c] >= total * 0.5 else None


def existing_gib_cols(data, header_row):
    if not header_row or header_row > len(data):
        return {}
    cols = {}
    for c, v in enumerate(data[header_row - 1], 1):
        code = gib_code(v)
        if code and code not in cols:
            cols[code] = c
    return cols


def detect(data, start=None):
    """start=(row, col) — пользователь сам кликнул первое назначение."""
    if start:
        col, first = start[1], start[0]
        last = range_from(data, col, first)
    else:
        found = best_purpose_column(data)
        if not found:
            return None
        col, first, last = found
    header = header_row_above(data, col, first)
    lay = Layout(col, first, last, header, amount_column(data, col, first, last))
    lay.gib_cols = existing_gib_cols(data, header)
    return lay


NAME_HEADERS = ("номгӯи корхонаҳо", "номгуи корхонахо", "основание", "наименование", "назначение")


def by_header(data):
    """Таблица ещё без назначений: столбец по заголовку «Номгӯи корхонаҳо» / «Основание»."""
    for r, row in enumerate(data[:60], 1):
        for c, v in enumerate(row, 1):
            if isinstance(v, str) and v.strip().lower().startswith(NAME_HEADERS):
                lay = Layout(c, r + 1, r + 1, r)
                lay.gib_cols = existing_gib_cols(data, r)
                return lay
    return None


def sheet_score(data):
    """Какой лист открыть: где есть назначения с ГИБ; лист с шапкой ГИБ (2111, 2211…) — в приоритете."""
    purposes = sum(1 for row in data for v in row if isinstance(v, str) and "ГИБ" in v.upper() and parse_purpose(v))
    has_header = any(sum(1 for v in row if gib_code(v)) >= 3 for row in data[:40])
    return purposes + (100000 if has_header else 0)
