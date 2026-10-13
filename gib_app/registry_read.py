"""Реестры платежей (выгрузки казначейства, сотни тысяч строк): узнать по шапке и быстро прочитать."""
from dataclasses import dataclass
from datetime import date, datetime

from utils import find_col, norm_text, to_number

COLS = {"doc": ("№ док", "номер док", "№док"), "date": ("дата док",), "payer": ("плательщик",),
        "payer_inn": ("инн плат",), "amount": ("сумма",), "rec": ("получатель",), "rec_inn": ("инн получ",),
        "purpose": ("основание", "назначение")}
REQUIRED = ("payer", "payer_inn", "amount", "rec", "rec_inn")
HEAD_ROWS = 20


@dataclass
class Payment:
    """Одна строка реестра (row — номер строки в Excel)."""
    row: int
    doc: str
    date: date | None
    payer: str
    payer_inn: str
    rec: str
    rec_inn: str
    amount: float
    purpose: str = ""


def registry_columns(head) -> dict | None:
    """Столбцы реестра по заголовкам (0-based) или None, если это не реестр."""
    heads = [norm_text(v) for v in head]
    cols = {k: find_col(heads, p) for k, p in COLS.items()}
    if not all(cols[k] for k in REQUIRED):
        return None
    return {k: (c - 1 if c else None) for k, c in cols.items()}


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
    pinn, rinn = inn_text(get("payer_inn")), inn_text(get("rec_inn"))
    payer, rec = _text(get("payer")), _text(get("rec"))
    if not (pinn or rinn or payer or rec):
        return None
    return Payment(row, _text(get("doc")), _date(get("date")), payer, pinn, rec, rinn,
                   round(to_number(get("amount")) or 0.0, 2), _text(get("purpose")))


def read_registry(path: str, progress=lambda done: None):
    """Платежи из первого листа-реестра файла. → (имя листа, всего строк, генератор Payment)
    или (None, 0, None). progress(строк_прочитано) вызывается каждые 50 000 строк."""
    from python_calamine import CalamineWorkbook

    wb = CalamineWorkbook.from_path(path)
    for name in wb.sheet_names:
        sheet = wb.get_sheet_by_name(name)
        start = (sheet.start or (0, 0))[0]       # строки до первой заполненной calamine пропускает
        rows = sheet.iter_rows()
        for i, values in enumerate(rows):
            cols = registry_columns(values)
            if cols:
                return name, sheet.height, _payments(rows, cols, start + i + 2, progress)
            if i >= HEAD_ROWS:
                break
    return None, 0, None


def _payments(rows, cols, first_row, progress):
    for i, values in enumerate(rows):
        p = to_payment(first_row + i, values, cols)
        if p:
            yield p
        if i % 50000 == 0:
            progress(i)
