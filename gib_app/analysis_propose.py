"""Перед глубоким анализом: что программа думает о каждом файле — листы (первые строки для окна выбора)
и догадка: какой лист, строка заголовков, реестр или список, где ИНН и какие столбцы суммировать.
Решает пользователь — догадка только подставляется в окно «Как читать файл»."""
from analysis_list_map import AS_LIST, _find
from analysis_lists import money_cols
from compare_finder import find_tables
from registry_columns import guess_header, registry_columns, two_row_columns

REGISTRY = "registry"
PREVIEW_ROWS = 30           # столько первых строк листа показывается в окне
FULL_READ_LIMIT = 200_000   # таблицы больше — без поиска ИНН по содержимому (долго)


def _sheet_rows(sheet) -> list:
    rows = []
    for i, values in enumerate(sheet.iter_rows()):
        rows.append(["" if v is None else v for v in values])
        if i + 1 >= PREVIEW_ROWS:
            break
    return rows


def auto_columns(rows: list, kind: str) -> dict:
    """Догадка по первым строкам листа: {"head", "cols", "sums", "found"} — для реестра или для списка
    (found — реестр узнан по заголовкам)."""
    if kind == REGISTRY:
        for i, values in enumerate(rows):
            cols = registry_columns(values) or (i > 0 and two_row_columns(rows[i - 1], values))
            if cols:
                return {"head": i, "cols": {k: v for k, v in cols.items() if k != "amount"}, "sums": [cols["amount"]],
                        "found": True}
        g = guess_header(rows)
        return {"head": g["head"], "cols": {k: v for k, v in g["cols"].items() if k != "amount"},
                "sums": [g["cols"]["amount"]] if g["cols"].get("amount") is not None else [], "found": False}
    g = guess_header(rows, _find)
    return {"head": g["head"], "cols": {"inn": g["cols"]["inn"], "name": g["cols"]["name"]},
            "sums": [g["cols"]["amount"]] if g["cols"]["amount"] is not None else [], "found": False}


def _list_guess(wb, name: str) -> dict | None:
    """Список: таблица с ИНН по содержимому (ИНН, название, суммы — как нашла бы программа)."""
    sheet = wb.get_sheet_by_name(name)
    data = [[None if v == "" else v for v in r] for r in sheet.to_python()]     # с первой заполненной строки
    found = find_tables(data)
    if not found:
        return None
    s = found[1][0]
    return {"sheet": name, "kind": AS_LIST, "head": (sheet.start or (0, 0))[0] + s.header - 1,
            "cols": {"inn": s.key_col - 1, "name": s.name_col - 1 if s.name_col else None},
            "sums": [c - 1 for c in money_cols(data, s)]}


def propose(path: str) -> dict:
    """→ {"path", "sheets": {лист: {"rows", "row0", "col0", "height"}}, "guess": {...}} (пустые листы пропущены)."""
    from python_calamine import CalamineWorkbook

    wb = CalamineWorkbook.from_path(path)
    sheets = {}
    for name in wb.sheet_names:
        sheet = wb.get_sheet_by_name(name)
        rows = _sheet_rows(sheet)
        if any(v != "" for r in rows for v in r):
            # iter_rows: строки — с первой строки листа (row0 = 0), столбцы — с первого заполненного (col0)
            sheets[name] = {"rows": rows, "row0": 0, "col0": (sheet.start or (0, 0))[1], "height": sheet.height}
    guess = None
    for name, sh in sheets.items():             # реестр — по заголовкам
        g = auto_columns(sh["rows"], REGISTRY)
        if g["found"]:
            guess = dict(g, sheet=name, kind=REGISTRY)
            break
    for name, sh in sheets.items():             # иначе — список с ИНН по содержимому
        if guess:
            break
        if sh["height"] <= FULL_READ_LIMIT:
            guess = _list_guess(wb, name)
    if not guess and sheets:
        name = next(iter(sheets))
        guess = dict(auto_columns(sheets[name]["rows"], AS_LIST), sheet=name, kind=AS_LIST)
    return {"path": path, "sheets": sheets, "guess": guess}


def propose_files(paths: list, progress=lambda p, t: None) -> list:
    out = []
    for k, path in enumerate(paths):
        progress(int(100 * k / len(paths)), f"Смотрю файл {k + 1} из {len(paths)}…")
        out.append(propose(path))
    return out
