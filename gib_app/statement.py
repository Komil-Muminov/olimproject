"""Выписка → платежи для переноса; поиск листа, куда их переносить (без изменения файлов)."""
from dataclasses import dataclass, field
from datetime import datetime

from constants import SUM_ROW_LABELS
from detect import by_header
from excel_reader import Book
from parser import parse_purpose_ex
from utils import norm_text, to_number

PURPOSE_HEADERS = ("основание", "назначение", "наименование")
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


def _col(heads, prefixes, after=0):
    return next((c for c, h in enumerate(heads, 1) if c > after and h.startswith(prefixes)), None)


def statement_header(data):
    """(строка шапки, назначение, расходы, № док, дата) — по заголовкам «Основание» и «Расходы»."""
    for r, row in enumerate(data[:60], 1):
        heads = [norm_text(v) for v in row]
        pc = _col(heads, PURPOSE_HEADERS)
        ac = _col(heads, ("расход",), pc or 0)
        if pc and ac:
            return r, pc, ac, _col(heads, ("№ док", "№док", "номер док", "№")), _col(heads, ("дата док", "дата"))
    return None


def read_statement(path: str):
    """Все платежи (расходы > 0) из выписки: (лист, [Payment])."""
    for name, data in Book(path).sheets.items():
        head = statement_header(data)
        if not head:
            continue
        hr, pc, ac, dc, tc = head
        out = []
        for r in range(hr + 1, len(data) + 1):
            text = Book.value(data, r, pc)
            amount = to_number(Book.value(data, r, ac))
            if not isinstance(text, str) or not text.strip() or not amount or amount <= 0:
                continue
            if norm_text(text).startswith(SUM_ROW_LABELS):      # «Обороты за период:», «Итого»
                continue
            gibs, longs = parse_purpose_ex(text)
            out.append(Payment(r, text.strip(), round(amount, 2),
                               Book.value(data, r, dc) if dc else None,
                               to_date(Book.value(data, r, tc)) if tc else None, gibs, longs))
        return name, out
    return None, []


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
