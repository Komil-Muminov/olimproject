"""Подсветка строк без ГИБ и сверка записанных сумм с итогами после заполнения."""
from constants import NO_GIB_NOTE, XL_NO_GIB, XL_NO_GIB_FONT
from excel_columns import col_values, gib_columns, used_bounds
from utils import to_num_safe


def mark_no_gib(ws, col: int, no_gib_rows: list, gib_rows: list):
    """Назначение без ГИБ — красная заливка, тёмно-красный текст и примечание.
    Если назначение исправили (ГИБ появился) — наша подсветка и примечание снимаются."""
    for r in no_gib_rows:
        c = ws.Cells(r, col)
        c.Interior.Color = XL_NO_GIB
        c.Font.Color = XL_NO_GIB_FONT
        try:
            c.ClearComments()
            c.AddComment(NO_GIB_NOTE)
        except Exception:
            pass
    for r in gib_rows:
        c = ws.Cells(r, col)
        if int(c.Interior.Color) == XL_NO_GIB:
            c.Interior.Pattern = -4142
            c.Font.ColorIndex = -4105
            try:
                if c.Comment is not None and c.Comment.Text() == NO_GIB_NOTE:
                    c.ClearComments()
            except Exception:
                pass


def verify_rows(ws, items: list, header_row: int, last_row: int) -> dict:
    """По каждому ГИБ: сколько должны были поставить, сколько стоит и сходится ли итог «Чамъ»."""
    used_last, _ = used_bounds(ws)
    expected = {}
    for _, gibs in items:
        for code, amount in gibs.items():
            expected[code] = round(expected.get(code, 0.0) + amount, 2)

    result = []
    for code, col in sorted(gib_columns(ws, header_row).items()):
        values = col_values(ws, col, header_row + 1, last_row)
        written = round(sum(to_num_safe(values[r - header_row - 1]) for r, g in items if code in g), 2)
        column_sum = round(sum(to_num_safe(v) for v in values), 2)
        total = None
        if used_last > last_row:
            formulas = col_values(ws, col, last_row + 1, used_last, "Formula")
            for i, f in enumerate(formulas):
                if isinstance(f, str) and f.startswith("="):
                    total = round(to_num_safe(ws.Cells(last_row + 1 + i, col).Value), 2)
                    break
        exp = expected.get(code, 0.0)
        if not exp and not column_sum:
            continue
        result.append({
            "code": code,
            "expected": exp,
            "written": written,
            "column_sum": column_sum,
            "total": total,
            "written_ok": abs(written - exp) < 0.005,
            "total_ok": total is None or abs(total - column_sum) < 0.005,
        })
    return {
        "gibs": result,
        "written_sum": round(sum(r["written"] for r in result), 2),
        "ok": all(r["written_ok"] and r["total_ok"] for r in result),
    }
