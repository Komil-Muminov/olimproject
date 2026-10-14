"""Столбцы для «Заполнить ГИБ»: выписки и файла, куда заполнять. Если по заголовкам не узнаны —
догадка о шапке для окна «Укажите столбцы» (столбцы 0-based, строки — номер в data, с 0)."""
from excel_reader import Book
from registry_columns import guess_header
from utils import find_col, norm_text

PURPOSE = ("основание", "назначение", "наименование", "номгуи корхонахо")   # norm_text: ӯ→у, ҳ→х

STATEMENT_FIELDS = ("purpose", "amount", "doc", "date")
STATEMENT_TITLES = {"purpose": "Назначение (основание) платежа", "amount": "Сумма расхода",
                    "doc": "№ документа", "date": "Дата документа"}
# «Расходы» не спрашиваем: суммы всегда идут в столбец «Расходы» (найдётся или создастся) — на нём сверка
TARGET_FIELDS = ("purpose", "doc", "date", "num")
TARGET_TITLES = {"purpose": "Назначение (Номгӯи корхонаҳо)", "doc": "№ ҳуҷҷат", "date": "Сана (дата)",
                 "num": "р/т (номер по порядку)"}


def _zero(c):
    return c - 1 if c else None


def statement_cols(head) -> dict:
    """Столбцы выписки по заголовкам: назначение, сумма расхода (правее назначения), № док., дата."""
    heads = [norm_text(v) for v in head]
    pc = find_col(heads, PURPOSE)
    return {"purpose": _zero(pc), "amount": _zero(find_col(heads, ("расход",), pc or 0)),
            "doc": _zero(find_col(heads, ("№ док", "№док", "номер док"))),
            "date": _zero(find_col(heads, ("дата док", "дата", "сана")))}


def statement_ok(cols: dict) -> bool:
    return cols.get("purpose") is not None and cols.get("amount") is not None


def target_cols(head) -> dict:
    """Столбцы файла, куда заполнять (как в Бучете)."""
    heads = [norm_text(v) for v in head]
    pc = find_col(heads, PURPOSE)
    doc = next((c for c, h in enumerate(heads, 1) if "хуччат" in h or "№ док" in h), None)
    return {"purpose": _zero(pc), "rash": _zero(find_col(heads, ("расход",), pc or 0)), "doc": _zero(doc),
            "date": _zero(find_col(heads, ("сана", "дата"))),
            "num": _zero(find_col(heads, ("р/т", "р\\т", "т/р", "№ п/п", "п/п")))}


def target_ok(cols: dict) -> bool:
    return cols.get("purpose") is not None


def book_guess(path: str, finder) -> dict | None:
    """Самый похожий лист и строка шапки (для окна выбора столбцов)."""
    best = None
    for name, data in Book(path).sheets.items():
        if not data:
            continue
        g = dict(guess_header(data, finder), sheet=name, row0=0, col0=0)
        if best is None or g["score"] > best["score"]:
            best = g
    return best
