"""«Платежи этого ИНН — в Excel»: все платежи организации из реестров (где платила и где получала)."""
import os

from openpyxl import Workbook

from excel_com import WriteError
from registry_read import read_registry
from report_fast import table_sheet
from utils import format_money

HEAD = ["№", "Файл", "Строка", "№ док.", "Дата", "Плательщик", "ИНН плат.", "Получатель", "ИНН получ.", "Сумма",
        "Основание"]
WIDTHS = [7, 14, 9, 13, 12, 36, 12, 36, 12, 16, 60]


def clean_inn(text: str) -> str:
    """ИНН из введённого текста: без пробелов; потерянный ноль в начале возвращается (40000011 → 040000011)."""
    inn = "".join(str(text or "").split())
    return inn.zfill(9) if inn.isdigit() and len(inn) < 9 else inn


def payments_of(paths, inn: str, mappings=None, progress=lambda p, t: None) -> tuple:
    """→ ([(файл, Payment)] где ИНН — плательщик или получатель, сколько файлов-реестров прочитано)."""
    out, registries, mappings = [], 0, mappings or {}
    for k, path in enumerate(paths):
        progress(int(85 * k / len(paths)), f"Ищу в {os.path.basename(path)} ({k + 1} из {len(paths)})…")
        if mappings.get(path) == "list":            # пользователь сказал: это список организаций, не реестр
            continue
        try:
            sheets, _ = read_registry(path, mappings.get(path))
        except Exception:
            continue                                # не реестр (обычная таблица) — пропускаем
        registries += bool(sheets)
        for _, _, rows in sheets:
            out += [(path, p) for p in rows if inn in (p.payer_inn, p.rec_inn)]
    return out, registries


def write_inn_payments(path: str, paths: list, inn: str, mappings=None, progress=lambda p, t: None) -> dict:
    """Два листа: где ИНН платил и где получал. Ничего не нашлось — файл не создаётся (path=None)."""
    inn = clean_inn(inn)
    found, registries = payments_of(paths, inn, mappings, progress)
    if not registries:
        raise WriteError("Среди выбранных файлов нет реестров платежей (столбцы «Плательщик», «ИНН», «Получатель», "
                         "«Сумма»), поэтому искать платежи негде.")
    found.sort(key=lambda x: (x[1].date.toordinal() if x[1].date else 10 ** 7, x[1].row))   # без даты — в конце
    paid = [x for x in found if x[1].payer_inn == inn]
    got = [x for x in found if x[1].rec_inn == inn]
    res = {"path": None, "inn": inn, "paid": len(paid), "got": len(got),
           "paid_sum": round(sum(p.amount for _, p in paid), 2), "got_sum": round(sum(p.amount for _, p in got), 2)}
    if not found:
        return res
    progress(90, "Сохраняю…")
    rows = lambda side: [[i, os.path.basename(f), p.row, p.doc, p.date, p.payer, p.payer_inn, p.rec, p.rec_inn,
                          p.amount, p.purpose] for i, (f, p) in enumerate(side, 1)]
    wb = Workbook(write_only=True)
    for name, side, total, color, hint in (
            ("Платил", paid, res["paid_sum"], "0F766E", "Платежи, где этот ИНН — плательщик."),
            ("Получал", got, res["got_sum"], "2563EB", "Платежи, где этот ИНН — получатель.")):
        table_sheet(wb, name, f"ИНН {inn} · {name.upper()} на {format_money(total)} · платежей", color, hint, HEAD, WIDTHS,
                    rows(side), money=(10,), dates=(5,))
    wb.save(path)
    res["path"] = path
    return res
