"""Сверка с выпиской — как в «Банк+муҳосибӣ (Бучет) -ислох шуда.xlsm».

Проверочный столбец «Расходы» стоит сразу за «Ҳамагӣ» и в суммы ГИБ не входит.
Сумма из выписки в нём заменяется формулой баланса =ROUND(Ҳамагӣ − сумма; 2):
0 — суммы совпали, иначе ячейка красная. Сама сумма видна в строке формул.
"""
import re

from openpyxl.utils import column_index_from_string, get_column_letter

from constants import RED_FILL, RED_FONT, TOTAL_HEADERS
from excel_columns import col_values, gib_columns, insert_column, row_values, sum_row_below, used_bounds
from excel_reader import cell_name
from utils import norm_text, to_num_safe, to_number

SIMPLE_SUM_RE = re.compile(r"^=SUM\(\$?([A-Z]+)\$?\d+:\$?([A-Z]+)\$?\d+\)$", re.IGNORECASE)
BALANCE_PREFIX = "=ROUND("


def find_header(ws, header_row: int, after_col: int, prefixes: tuple) -> int | None:
    _, last_col = used_bounds(ws)
    for c, v in enumerate(row_values(ws, header_row, last_col), 1):
        if c > after_col and norm_text(v).startswith(prefixes):
            return c
    return None


def _num_literal(v: float) -> str:
    """Число для формулы: с точкой, отрицательное — в скобках (как NumLiteral в макросе)."""
    s = f"{round(v, 2):.2f}".rstrip("0").rstrip(".")
    return f"({s})" if v < 0 else s


def setup_reconciliation(ws, header_row: int, purpose_col: int, first_row: int, last_row: int | None) -> dict:
    """Порядок столбцов: … | Ҳамагӣ | Расходы | ГИБ… Нет «Ҳамагӣ» — создаётся (в выписке —
    перед её столбцом «Расходы»); нет «Расходы» сразу за «Ҳамагӣ» — создаётся."""
    total = find_header(ws, header_row, purpose_col, TOTAL_HEADERS)
    total_created = total is None
    if total is None:
        rash_any = find_header(ws, header_row, purpose_col, ("расход",))
        if rash_any:
            total = rash_any
        else:
            _, last_col = used_bounds(ws)
            vals = row_values(ws, header_row, last_col)
            total = max((i for i, v in enumerate(vals, 1) if v not in (None, "")), default=purpose_col) + 1
        insert_column(ws, header_row, total, "Ҳамагӣ")
    rash = total + 1
    if not norm_text(ws.Cells(header_row, rash).Value).startswith("расход"):
        insert_column(ws, header_row, rash, "Расходы")

    if last_row is None:
        sr = sum_row_below(ws, first_row, purpose_col)
        last_row = sr - 1 if sr else first_row
    sum_row = sum_row_below(ws, last_row, purpose_col)
    if sum_row:
        last_row = sum_row - 1          # вся таблица до «Чамъ», а не только заполненные строки
    set_balance_formulas(ws, total, rash, first_row, last_row)

    rng = ws.Range(ws.Cells(first_row, rash), ws.Cells(last_row, rash))
    try:
        fc = rng.FormatConditions.Add(Type=1, Operator=4, Formula1="=0")     # баланс ≠ 0 — красный
        fc.Interior.Color = RED_FILL
        fc.Font.Color = RED_FONT
        fc.Font.Bold = True
    except Exception:
        pass
    if sum_row:
        for col in (rash,) + ((total,) if total_created else ()):
            cell = ws.Cells(sum_row, col)
            if cell.Formula in ("", None) or SIMPLE_SUM_RE.match(str(cell.Formula)):
                cell.Formula = f"=SUM({cell_name(first_row, col)}:{cell_name(last_row, col)})"
    return {"total": total, "rash": rash, "total_created": total_created, "sum_row": sum_row}


def set_balance_formulas(ws, total: int, rash: int, first_row: int, last_row: int):
    """Суммы (числа) в «Расходы» → =ROUND(Ҳамагӣ − сумма; 2). Формулы и пустые ячейки не трогаются."""
    formulas = col_values(ws, rash, first_row, last_row, "Formula")
    values = col_values(ws, rash, first_row, last_row)
    for i, (f, v) in enumerate(zip(formulas, values)):
        if isinstance(f, str) and f.startswith("="):
            continue
        amt = to_number(v)
        if amt is not None:
            r = first_row + i
            ws.Cells(r, rash).Formula = f"=ROUND({cell_name(r, total)}-{_num_literal(amt)},2)"


def fill_total_formulas(ws, header_row: int, total: int, first_row: int, last_row: int, rash: int | None = None):
    """«Ҳамагӣ» = сумма столбцов ГИБ строки; «Расходы» в сумму не входит."""
    cols = sorted(gib_columns(ws, header_row).values())
    if not cols:
        return
    rng = ws.Range(ws.Cells(first_row, total), ws.Cells(last_row, total))
    out = []
    for r, f in zip(range(first_row, last_row + 1), col_values(ws, total, first_row, last_row, "Formula")):
        a, b = cols[0], cols[-1]
        m = SIMPLE_SUM_RE.match(str(f or ""))
        if f not in ("", None) and not m:
            out.append(f)
            continue
        if m:
            start = column_index_from_string(m.group(1).upper())
            if not rash or start > rash:
                a = min(a, start)
            b = max(b, column_index_from_string(m.group(2).upper()))
        out.append(f"=SUM({get_column_letter(a)}{r}:{get_column_letter(b)}{r})")
    rng.Formula = [(v,) for v in out]


def read_reconciliation(ws, header_row: int, first_row: int, last_row: int, rash: int | None = None):
    """Итог сверки по столбцу «Расходы»: строки с балансом ≠ 0 и есть ли вообще суммы из выписки."""
    rash = rash or find_header(ws, header_row, 0, ("расход",))
    if not rash:
        return None
    formulas = col_values(ws, rash, first_row, last_row, "Formula")
    values = col_values(ws, rash, first_row, last_row)
    balances = [(first_row + i, to_num_safe(v)) for i, (f, v) in enumerate(zip(formulas, values))
                if isinstance(f, str) and f.upper().startswith(BALANCE_PREFIX)]
    bad = [r for r, v in balances if abs(v) > 0.005]
    return {"empty": not balances, "bad_rows": bad, "total_diff": round(sum(v for _, v in balances), 2)}


NOTE_PREFIX = "По назначению"


def _money(v: float) -> str:
    return f"{v:,.2f}".replace(",", " ").replace(".", ",")


def update_balance_notes(ws, total: int, rash: int, first_row: int, last_row: int):
    """Примечание у ячейки «Расходы», где баланс не 0: откуда взялась разница.
    Вызывать после пересчёта книги. Совпало — наше примечание убирается."""
    formulas = col_values(ws, rash, first_row, last_row, "Formula")
    balances = col_values(ws, rash, first_row, last_row)
    totals = col_values(ws, total, first_row, last_row)
    for i, (f, bal, tot) in enumerate(zip(formulas, balances, totals)):
        c = ws.Cells(first_row + i, rash)
        is_balance = isinstance(f, str) and f.upper().startswith(BALANCE_PREFIX)
        bal, tot = to_num_safe(bal), to_num_safe(tot)
        try:
            ours = c.Comment is not None and str(c.Comment.Text()).startswith(NOTE_PREFIX)
            if is_balance and abs(bal) > 0.004:
                if c.Comment is not None:
                    c.Comment.Delete()
                amt = round(tot - bal, 2)
                side = "меньше" if bal < 0 else "больше"
                c.AddComment(f"{NOTE_PREFIX} (Ҳамагӣ): {_money(tot)}\nПо выписке: {_money(amt)}\n"
                             f"Разница: {_money(bal)} — в назначении на {_money(abs(bal))} {side}")
                c.Comment.Shape.TextFrame.AutoSize = True
            elif ours:
                c.Comment.Delete()
        except Exception:
            pass
