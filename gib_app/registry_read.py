"""Реестры платежей (выгрузки казначейства, сотни тысяч строк): узнать по шапке и быстро прочитать."""
from dataclasses import dataclass
from datetime import date, datetime

from registry_columns import HEAD_ROWS, guess_header, registry_columns, two_row_columns
from utils import to_number


@dataclass
class Payment:
    """Одна строка реестра (row — номер строки в Excel)."""
    row: int
    doc: str
    date: date | None
    payer: str
    payer_inn: str | None       # None — в документе нет такого столбца (только одна сторона)
    rec: str
    rec_inn: str | None
    amount: float
    purpose: str = ""


def inn_text(v) -> str:
    """ИНН текстом; числом записанный ИНН дополняется нулями слева (10002120 → 010002120)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, int):
        return str(v).zfill(9)
    return str(v).strip()


def _text(v) -> str:
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip() if v is not None else ""


def _date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip()[:10], fmt).date()
            except ValueError:
                pass
    return None


def to_payment(row: int, values, cols: dict) -> Payment | None:
    """Строка → Payment; строки итогов (без обоих ИНН и названий) пропускаются."""
    def get(k):
        c = cols.get(k)
        return values[c] if c is not None and c < len(values) else None
    side = lambda k: inn_text(get(k)) if cols.get(k) is not None else None
    pinn, rinn = side("payer_inn"), side("rec_inn")
    payer, rec = _text(get("payer")), _text(get("rec"))
    if not (pinn or rinn or payer or rec):
        return None
    money = cols.get("amount")         # список — несколько столбцов, выбранных пользователем (суммируются)
    if isinstance(money, list):
        amount = sum(to_number(values[c]) or 0.0 for c in money if c < len(values))
    else:
        amount = to_number(get("amount")) or 0.0
    return Payment(row, _text(get("doc")), _date(get("date")), payer, pinn, rec, rinn, round(amount, 2),
                   _text(get("purpose")))


def read_registry(path: str, mapping: dict | None = None):
    """Все листы-реестры файла. → (листы [(имя, всего строк, генератор Payment)], догадка).
    Если ни один лист не узнан, догадка — самый похожий на реестр лист и его шапка (для окна выбора
    столбцов): {"sheet", "head", "heads", "cols", "score"}. mapping — столбцы, указанные пользователем:
    {"sheet", "head", "cols"}."""
    from python_calamine import CalamineWorkbook

    wb = CalamineWorkbook.from_path(path)
    sheets, guess = [], None
    for name in wb.sheet_names:
        if mapping and name != mapping["sheet"]:
            continue
        sheet = wb.get_sheet_by_name(name)
        rows = sheet.iter_rows()
        head = []
        for i, values in enumerate(rows):
            head.append(values)
            if mapping:
                cols = mapping["cols"] if i == mapping["head"] else None
            else:
                cols = registry_columns(values) or (i > 0 and two_row_columns(head[-2], values))
            if cols:
                sheets.append((name, sheet.height, _payments(rows, cols, i + 2)))
                break
            if i >= HEAD_ROWS and not mapping:      # строку заголовков указал пользователь — ищем её дальше
                break
        if not sheets and head:
            # iter_rows: строки — с первой строки листа, столбцы — с первого заполненного (col0)
            g = dict(guess_header(head), sheet=name, row0=0, col0=(sheet.start or (0, 0))[1])
            if guess is None or g["score"] > guess["score"]:
                guess = g
    return sheets, (None if sheets else guess)


def _payments(rows, cols, first_row):
    for i, values in enumerate(rows):
        p = to_payment(first_row + i, values, cols)
        if p:
            yield p
