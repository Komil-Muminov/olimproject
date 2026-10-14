"""Пометка «Зафиксировать распределение» из макроса Бучета (как в новой версии файла).

Зафиксированная сумма по статье — ячейка с форматом FROZEN_FMT (выглядит как «0.00»). Новые строки
копируют оформление строки выше и могут унаследовать пометку — тогда макрос посчитал бы их
зафиксированными и не пересчитывал. Программа снимает такую пометку в строках, которые пишет сама."""
from excel_columns import gib_columns
from excel_com import number_format, set_number_format

FROZEN_FMT = "0.00;-0.00;0.00;@"
AMT_FMT = "0.00"
OLD_FROZEN_FMT = "General;-General;General;@"     # прежняя версия: пометка в ячейке назначения
NAME_FMT = "@"


def _fix(ws, col: int, r1: int, r2: int, mark: str, plain: str):
    rng = ws.Range(ws.Cells(r1, col), ws.Cells(r2, col))
    fmt = number_format(rng)
    if fmt == mark:
        set_number_format(rng, plain)
    elif fmt is None:                       # разные форматы в столбце — по ячейкам
        for r in range(r1, r2 + 1):
            if number_format(ws.Cells(r, col)) == mark:
                set_number_format(ws.Cells(r, col), plain)


def clear_freeze_marks(ws, header_row: int, purpose_col: int, r1: int, r2: int):
    """Строки r1..r2 — новые, не зафиксированные: снять унаследованные пометки фиксации."""
    if r2 < r1:
        return
    for col in gib_columns(ws, header_row).values():
        _fix(ws, col, r1, r2, FROZEN_FMT, AMT_FMT)
    _fix(ws, purpose_col, r1, r2, OLD_FROZEN_FMT, NAME_FMT)
