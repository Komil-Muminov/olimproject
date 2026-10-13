"""Таблица в файле, куда переносится выписка (через Excel COM): столбцы, свободные строки, итоги."""
import re
from collections import Counter
from dataclasses import dataclass

from openpyxl.utils import get_column_letter

from constants import XL_DOWN
from excel_columns import col_values, row_values, sum_row_below, used_bounds
from excel_reconcile import SIMPLE_SUM_RE
from statement import doc_text
from utils import norm_text, to_number

BALANCE_AMOUNT_RE = re.compile(r"-\(?(-?[\d.]+)\)?,2\)$")


@dataclass
class Target:
    header_row: int
    purpose: int
    first_row: int
    last_filled: int
    sum_row: int | None
    num: int | None = None
    doc: int | None = None
    date: int | None = None
    rash: int | None = None


def _col(heads, prefixes):
    return next((c for c, h in enumerate(heads, 1) if h.startswith(prefixes)), None)


def locate(ws, header_row: int, purpose_col: int) -> Target:
    _, last_col = used_bounds(ws)
    heads = [norm_text(v) for v in row_values(ws, header_row, last_col)]
    t = Target(header_row, purpose_col, header_row + 1, header_row, sum_row_below(ws, header_row, purpose_col))
    t.num = _col(heads, ("р/т", "р\\т", "т/р", "№ п/п", "п/п"))
    t.doc = next((c for c, h in enumerate(heads, 1) if "хуччат" in h or "№ док" in h), None)
    t.date = _col(heads, ("сана", "дата"))
    t.rash = next((c for c, h in enumerate(heads, 1) if c > purpose_col and h.startswith("расход")), None)
    end = (t.sum_row - 1) if t.sum_row else max(used_bounds(ws)[0], header_row + 1)
    texts = col_values(ws, purpose_col, header_row + 1, end)
    nums = col_values(ws, t.num, header_row + 1, end) if t.num else [None] * len(texts)
    # первая строка данных: с номером р/т или с назначением (строка под шапкой может быть пустой)
    for i, (txt, n) in enumerate(zip(texts[:5], nums[:5])):
        if to_number(n) is not None or str(txt or "").strip():
            t.first_row = header_row + 1 + i
            break
    filled = [header_row + 1 + i for i, v in enumerate(texts) if str(v or "").strip()]
    t.last_filled = max(filled) if filled else t.first_row - 1
    return t


def _amount(formula, value):
    """Сумма из выписки в «Расходы»: число или =ROUND(Ҳамагӣ-сумма,2)."""
    m = BALANCE_AMOUNT_RE.search(str(formula or ""))
    return round(float(m.group(1)), 2) if m else (round(to_number(value), 2) if to_number(value) is not None else None)


def payment_key(purpose, amount, doc=None):
    return norm_text(purpose), amount, doc_text(doc)


def existing_keys(ws, t: Target, use_doc: bool) -> Counter:
    """Что уже перенесено раньше: назначение + сумма (+ № документа)."""
    if t.last_filled < t.first_row:
        return Counter()
    r1, r2 = t.first_row, t.last_filled
    texts = col_values(ws, t.purpose, r1, r2)
    forms = col_values(ws, t.rash, r1, r2, "Formula") if t.rash else [None] * len(texts)
    vals = col_values(ws, t.rash, r1, r2) if t.rash else [None] * len(texts)
    docs = col_values(ws, t.doc, r1, r2) if use_doc else [None] * len(texts)
    return Counter(payment_key(p, _amount(f, v), d) for p, f, v, d in zip(texts, forms, vals, docs)
                   if str(p or "").strip())


def make_room(ws, t: Target, count: int) -> int:
    """Свободные строки сразу после последней заполненной; не хватает до «Чамъ» — вставляются.
    Возвращает первую строку для записи."""
    start = t.last_filled + 1
    if not t.sum_row:
        return start
    need = start + count - t.sum_row
    if need > 0:
        ws.Range(ws.Rows(t.sum_row), ws.Rows(t.sum_row + need - 1)).Insert(Shift=XL_DOWN, CopyOrigin=0)
        t.sum_row += need
        extend_sum_row(ws, t)
    return start


NEW_HEADERS = ("р/т", "Номгӯи корхонаҳо", "№ ҳуҷҷат", "Сана", "Ҳамагӣ", "Расходы")
NEW_WIDTHS = (6, 60, 12, 12, 14, 14)


def create_table(ws) -> tuple:
    """Пустой лист → таблица как в Бучете: шапка (строка 1), одна пустая строка, «Чамъ».
    Возвращает (строка шапки, столбец назначений)."""
    n = len(NEW_HEADERS)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n)).Value = [NEW_HEADERS]
    ws.Cells(3, 1).Value = "Чамъ"
    for c in (5, 6):
        L = get_column_letter(c)
        ws.Cells(3, c).Formula = f"=SUM({L}2:{L}2)"
    for c, w in enumerate(NEW_WIDTHS, 1):
        ws.Columns(c).ColumnWidth = w
    ws.Columns(2).WrapText = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n)).Font.Bold = True
    ws.Range(ws.Cells(3, 1), ws.Cells(3, n)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(3, n)).Borders.LineStyle = 1
    return 1, 2


def extend_sum_row(ws, t: Target):
    """Итоги «Чамъ» вида =SUM(G6:G9) растягиваются на всю таблицу (с новыми строками)."""
    _, last_col = used_bounds(ws)
    forms = ws.Range(ws.Cells(t.sum_row, 1), ws.Cells(t.sum_row, last_col)).Formula
    for c, f in enumerate(forms[0] if isinstance(forms, tuple) else (forms,), 1):
        m = SIMPLE_SUM_RE.match(str(f or ""))
        if m and m.group(1).upper() == m.group(2).upper():
            L = get_column_letter(c)
            ws.Cells(t.sum_row, c).Formula = f"=SUM({L}{t.first_row}:{L}{t.sum_row - 1})"
