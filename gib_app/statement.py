"""Выписка → платежи для переноса; поиск листа, куда их переносить (без изменения файлов)."""
from dataclasses import dataclass, field
from datetime import datetime

from constants import SUM_ROW_LABELS
from detect import by_header
from excel_reader import Book
from parser import parse_purpose_ex
from gib_columns import statement_cols, statement_ok
from utils import norm_text, to_number

BUDGET_WORDS = ("буҷет", "бучет", "бюджет", "budget")


@dataclass
class Payment:
    row: int
    purpose: str
    amount: float
    doc: object = None
    date: object = None
    gibs: dict = field(default_factory=dict)
    longs: list = field(default_factory=list)


def to_date(v):
    """«22.12.2025» → datetime; дата Excel остаётся датой; прочее — как есть."""
    if isinstance(v, datetime) or v in (None, ""):
        return v
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(v).strip()[:10], fmt)
        except ValueError:
            pass
    return v


def doc_text(v) -> str:
    """Номер документа без «.0» и пробелов — для сравнения."""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v or "").strip()
    return s.lstrip("0") or s      # Excel мог срезать нули: 0002309 = 2309


def statement_header(data):
    """(строка шапки с 1, столбцы с 0) — по заголовкам «Основание» и «Расходы»; не нашлось — None."""
    for r, row in enumerate(data[:60], 1):
        cols = statement_cols(row)
        if statement_ok(cols):
            return r, cols
    return None


def read_statement(path: str, mapping: dict | None = None):
    """Все платежи (расходы > 0) из выписки: (лист, [Payment]); шапка не найдена — (None, []).
    mapping — столбцы, указанные пользователем: {"sheet", "head" (с 0), "cols" (с 0)}."""
    for name, data in Book(path).sheets.items():
        if mapping:
            if name != mapping["sheet"]:
                continue
            hr, cols = mapping["head"] + 1, mapping["cols"]
        else:
            head = statement_header(data)
            if not head:
                continue
            hr, cols = head
        return name, _payments(data, hr, cols)
    return None, []


def _payments(data, hr: int, cols: dict) -> list:
    get = lambda r, k: Book.value(data, r, cols[k] + 1) if cols.get(k) is not None else None
    out = []
    for r in range(hr + 1, len(data) + 1):
        text, amount = get(r, "purpose"), to_number(get(r, "amount"))
        if not isinstance(text, str) or not text.strip() or not amount or amount <= 0:
            continue
        if norm_text(text).startswith(SUM_ROW_LABELS):      # «Обороты за период:», «Итого»
            continue
        gibs, longs = parse_purpose_ex(text)
        out.append(Payment(r, text.strip(), round(amount, 2), get(r, "doc"), to_date(get(r, "date")), gibs, longs))
    return out


def target_sheets(path: str) -> list:
    """Листы, куда можно перенести: [(имя, Layout)], подходящие — первыми
    (есть столбцы ГИБ, «Расходы», в названии «буҷет»). Таблицы нет — пустые листы с Layout=None:
    там программа создаст таблицу сама."""
    found, empty = [], []
    for i, (name, data) in enumerate(Book(path).sheets.items()):
        lay = by_header(data)
        if not lay:
            if not any(v not in (None, "") for row in data for v in row):
                empty.append((name, None))
            continue
        heads = [norm_text(v) for v in data[lay.header_row - 1]]
        score = (100 if lay.gib_cols else 0) + (50 if any(h.startswith("расход") for h in heads) else 0)
        score += 20 if any(w in name.lower() for w in BUDGET_WORDS) else 0
        found.append((-score, i, name, lay))
    return [(name, lay) for _, _, name, lay in sorted(found, key=lambda x: x[:2])] or empty
