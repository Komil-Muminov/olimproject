"""Excel со всеми платежами одного ИНН (поиск по ИНН в реестрах)."""
import os

from openpyxl import Workbook

from registry_read import read_registry
from report_fast import table_sheet

HEAD = ["№", "Файл", "Строка", "№ док.", "Дата", "Плательщик", "ИНН плат.", "Получатель", "ИНН получ.", "Сумма",
        "Основание"]
WIDTHS = [7, 14, 9, 13, 12, 36, 12, 36, 12, 16, 60]


def payments_of(paths, inn: str, mappings=None, progress=lambda p, t: None) -> list:
    """Все платежи, где ИНН — плательщик или получатель: [(файл, Payment)]."""
    out = []
    for k, path in enumerate(paths):
        progress(int(80 * k / len(paths)), f"Ищу в {os.path.basename(path)}…")
        if (mappings or {}).get(path) == "list":
            continue
        sheets, _ = read_registry(path, (mappings or {}).get(path))
        for _, _, rows in sheets:
            out += [(path, p) for p in rows if inn in (p.payer_inn, p.rec_inn)]
    return out


def write_inn_payments(path: str, paths: list, inn: str, mappings=None, progress=lambda p, t: None) -> dict:
    """Два листа: где ИНН платил и где получал. → {"path", "paid", "got"}."""
    by_date = lambda x: (x[1].date.toordinal() if x[1].date else 10 ** 7, x[1].row)      # без даты — в конце
    found = sorted(payments_of(paths, inn, mappings, progress), key=by_date)
    progress(90, "Сохраняю…")
    rows = lambda side: [[i, os.path.basename(f), p.row, p.doc, p.date, p.payer, p.payer_inn, p.rec, p.rec_inn,
                          p.amount, p.purpose] for i, (f, p) in enumerate(side, 1)]
    paid = [x for x in found if x[1].payer_inn == inn]
    got = [x for x in found if x[1].rec_inn == inn]
    wb = Workbook(write_only=True)
    table_sheet(wb, "Платил", f"ИНН {inn} — ПЛАТИЛ", "0F766E", "Платежи, где этот ИНН — плательщик.",
                HEAD, WIDTHS, rows(paid), money=(10,), dates=(5,))
    table_sheet(wb, "Получал", f"ИНН {inn} — ПОЛУЧАЛ", "2563EB", "Платежи, где этот ИНН — получатель.",
                HEAD, WIDTHS, rows(got), money=(10,), dates=(5,))
    wb.save(path)
    return {"path": path, "paid": len(paid), "got": len(got)}
