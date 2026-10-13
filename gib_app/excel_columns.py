"""Операции над столбцами и диапазонами рабочего листа через Excel COM."""
from constants import (
    SUM_ROW_LABELS, XL_FORMAT_FROM_LEFT_OR_ABOVE, XL_TO_RIGHT
)
from excel_reader import gib_code
from utils import norm_text


def used_bounds(ws):
    ur = ws.UsedRange
    return ur.Row + ur.Rows.Count - 1, ur.Column + ur.Columns.Count - 1


def row_values(ws, row: int, last_col: int):
    v = ws.Range(ws.Cells(row, 1), ws.Cells(row, last_col)).Value
    return list(v[0]) if isinstance(v, tuple) else [v]


def col_values(ws, col: int, r1: int, r2: int, prop="Value"):
    v = getattr(ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)), prop)
    return [x[0] for x in v] if isinstance(v, tuple) else [v]


def set_column(ws, row: int, col: int, values: list):
    ws.Range(ws.Cells(row, col), ws.Cells(row + len(values) - 1, col)).Value = [(v,) for v in values]


def gib_columns(ws, header_row: int) -> dict:
    _, last_col = used_bounds(ws)
    cols = {}
    for i, v in enumerate(row_values(ws, header_row, last_col), start=1):
        code = gib_code(v)
        if code and code not in cols:
            cols[code] = i
    return cols


def insert_column(ws, header_row: int, pos: int, title: str):
    ws.Columns(pos).Insert(Shift=XL_TO_RIGHT, CopyOrigin=XL_FORMAT_FROM_LEFT_OR_ABOVE)
    ws.Cells(header_row, pos).Value = title
    ws.Columns(pos).ColumnWidth = max(ws.Columns(pos).ColumnWidth, 13)


def sum_row_below(ws, last_row: int, purpose_col: int):
    """Строка итогов под таблицей: «Чамъ», «Итого», «Обороты за период»…"""
    used_last, _ = used_bounds(ws)
    if used_last <= last_row:
        return None
    block = ws.Range(ws.Cells(last_row + 1, 1), ws.Cells(used_last, purpose_col)).Value
    rows = block if isinstance(block, tuple) else ((block,),)
    for i, row in enumerate(rows):
        if any(norm_text(v).startswith(SUM_ROW_LABELS) for v in row):
            return last_row + 1 + i
    return None


def has_formulas(ws, col: int) -> bool:
    last_row, _ = used_bounds(ws)
    return any(isinstance(f, str) and f.startswith("=") for f in col_values(ws, col, 1, last_row, "Formula"))


def copy_column_formulas(ws, src: int, dst: int, header_row: int):
    """Формулы соседнего столбца ГИБ (например, итог «Ҷамъ») копируем в новый столбец."""
    last_row, _ = used_bounds(ws)
    formulas = col_values(ws, src, 1, last_row, "FormulaR1C1")
    for r, f in enumerate(formulas, start=1):
        if r != header_row and isinstance(f, str) and f.startswith("="):
            ws.Cells(r, dst).FormulaR1C1 = f


def swap_constants(ws, a: int, b: int):
    """Меняет местами значения двух столбцов (кроме формул)."""
    last_row, _ = used_bounds(ws)
    fa = col_values(ws, a, 1, last_row, "Formula")
    fb = col_values(ws, b, 1, last_row, "Formula")
    va = col_values(ws, a, 1, last_row)
    vb = col_values(ws, b, 1, last_row)
    for r in range(1, last_row + 1):
        is_fa = isinstance(fa[r - 1], str) and fa[r - 1].startswith("=")
        is_fb = isinstance(fb[r - 1], str) and fb[r - 1].startswith("=")
        if is_fa or is_fb or va[r - 1] == vb[r - 1]:
            continue
        ws.Cells(r, a).Value = vb[r - 1]
        ws.Cells(r, b).Value = va[r - 1]


def ensure_gib_columns(ws, header_row: int, needed: list, pcol: int) -> list:
    """Создаёт недостающие столбцы ГИБ по возрастанию кода. Возвращает список созданных."""
    created = []
    if needed and ws.Cells(header_row, pcol).Value in (None, ""):
        ws.Cells(header_row, pcol).Value = "Номгӯи корхонаҳо"
    for code in needed:
        cols = gib_columns(ws, header_row)
        if code in cols:
            continue
        created.append(code)
        if not cols:
            _, last_col = used_bounds(ws)
            vals = row_values(ws, header_row, last_col)
            last = max((i for i, v in enumerate(vals, 1) if v not in (None, "")), default=0)
            ws.Cells(header_row, max(last, pcol) + 1).Value = int(code)
            continue
        last_gib_col = max(cols.values())
        header_val = ws.Cells(header_row, last_gib_col + 1).Value
        if max(cols, key=int) < code and header_val in (None, "") and not has_formulas(ws, last_gib_col):
            ws.Cells(header_row, last_gib_col + 1).Value = int(code)
            continue

        ordered = sorted(cols.items(), key=lambda kv: kv[1])
        first_col, last_gib_col = ordered[0][1], ordered[-1][1]
        bigger = [c for k, c in ordered if k > code]
        swap = None
        if bigger and bigger[0] != first_col:
            pos = bigger[0]
        elif bigger:
            pos, swap = first_col + 1, first_col
        elif ws.Cells(header_row, last_gib_col + 1).Value not in (None, ""):
            pos = last_gib_col + 1
        elif len(ordered) > 1:
            pos, swap = last_gib_col, last_gib_col + 1
        else:
            pos = last_gib_col + 1

        ws.Columns(pos).Insert(Shift=XL_TO_RIGHT, CopyOrigin=XL_FORMAT_FROM_LEFT_OR_ABOVE)
        neighbour = pos - 1 if pos - 1 >= first_col else pos + 1
        copy_column_formulas(ws, neighbour, pos, header_row)
        ws.Cells(header_row, pos).Value = int(code)
        if swap is not None:
            swap_constants(ws, pos, swap)
    return created
