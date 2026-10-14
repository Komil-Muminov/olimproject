"""«Всё по ИНН — в Excel»: где встречается организация во всех выбранных файлах — строки обычных таблиц
(с подписями столбцов) и платежи из реестров (где платила и где получала)."""
import os

from openpyxl import Workbook

from analysis_list_map import as_list
from registry_read import read_registry
from report_fast import table_sheet
from report_inn_tables import tables_report_rows
from utils import format_money

HEAD = ["№", "Файл", "Строка", "№ док.", "Дата", "Плательщик", "ИНН плат.", "Получатель", "ИНН получ.", "Сумма",
        "Основание"]
WIDTHS = [7, 14, 9, 13, 12, 36, 12, 36, 12, 16, 60]
WHERE_HEAD = ["№", "Файл", "Лист", "Строка", "Что", "Название", "Сколько", "Сумма"]
WHERE_WIDTHS = [6, 30, 24, 9, 12, 44, 10, 18]
TABLE_HEAD = ["№", "Файл", "Лист", "Строка", "Столбец", "Значение"]
TABLE_WIDTHS = [7, 30, 24, 9, 50, 24]


def clean_inn(text: str) -> str:
    """ИНН из введённого текста: без пробелов; потерянный ноль в начале возвращается (40000011 → 040000011)."""
    inn = "".join(str(text or "").split())
    return inn.zfill(9) if inn.isdigit() and len(inn) < 9 else inn


def payments_of(paths, inn: str, mappings=None, progress=lambda p, t: None) -> tuple:
    """→ ([(файл, Payment)] где ИНН — плательщик или получатель, {пути реестров})."""
    out, registries, mappings = [], set(), mappings or {}
    for k, path in enumerate(paths):
        progress(int(60 * k / len(paths)), f"Ищу в {os.path.basename(path)} ({k + 1} из {len(paths)})…")
        if as_list(mappings.get(path)):             # пользователь сказал: это список организаций, не реестр
            continue
        try:
            sheets, _ = read_registry(path, mappings.get(path))
        except Exception:
            continue                                # не реестр (обычная таблица) — пропускаем
        if sheets:
            registries.add(path)
        for _, _, rows in sheets:
            out += [(path, p) for p in rows if inn in (p.payer_inn, p.rec_inn)]
    return out, registries


def _where_payments(side, what: str) -> list:
    """«Где нашлось»: по реестру — одна строка на файл (сколько платежей и на какую сумму)."""
    by = {}
    for f, p in side:
        n, s, name = by.get(f, (0, 0.0, ""))
        by[f] = (n + 1, s + p.amount, name or (p.payer if what == "платил" else p.rec))
    return [[os.path.basename(f), "", None, what, name, n, round(s, 2)] for f, (n, s, name) in by.items()]


def write_inn_payments(path: str, paths: list, inn: str, mappings=None, progress=lambda p, t: None) -> dict:
    """Листы: «Где нашлось», «Из таблиц» (строки обычных таблиц), «Платил» / «Получал» (из реестров).
    Ничего не нашлось — файл не создаётся (path=None)."""
    inn = clean_inn(inn)
    found, registries = payments_of(paths, inn, mappings, progress)
    table_rows, where = tables_report_rows(paths, inn, registries, progress)
    found.sort(key=lambda x: (x[1].date.toordinal() if x[1].date else 10 ** 7, x[1].row))   # без даты — в конце
    paid = [x for x in found if x[1].payer_inn == inn]
    got = [x for x in found if x[1].rec_inn == inn]
    res = {"path": None, "inn": inn, "paid": len(paid), "got": len(got), "tables": len(where),
           "registries": len(registries),
           "paid_sum": round(sum(p.amount for _, p in paid), 2), "got_sum": round(sum(p.amount for _, p in got), 2)}
    if not found and not where:
        return res
    progress(90, "Сохраняю…")
    rows = lambda side: [[i, os.path.basename(f), p.row, p.doc, p.date, p.payer, p.payer_inn, p.rec, p.rec_inn,
                          p.amount, p.purpose] for i, (f, p) in enumerate(side, 1)]
    wb = Workbook(write_only=True)
    where += _where_payments(paid, "платил") + _where_payments(got, "получал")
    table_sheet(wb, "Где нашлось", f"ИНН {inn} · ГДЕ НАШЛОСЬ", "7C3AED",
                "Все места во всех выбранных файлах. «Сколько» — у таблицы: заполненных ячеек в строке; "
                "у реестра: платежей.", WHERE_HEAD, WHERE_WIDTHS, [[i] + r for i, r in enumerate(where, 1)], money=(8,))
    if table_rows:
        table_sheet(wb, "Из таблиц", f"ИНН {inn} · ИЗ ТАБЛИЦ · значений", "0891B2",
                    "Строки обычных таблиц, где есть этот ИНН: каждое значение — с файлом, листом, строкой и "
                    "столбцом (буква и заголовок).", TABLE_HEAD, TABLE_WIDTHS, table_rows, money=(6,))
    sides = [("Платил", paid, res["paid_sum"], "0F766E", "Платежи, где этот ИНН — плательщик."),
             ("Получал", got, res["got_sum"], "2563EB", "Платежи, где этот ИНН — получатель.")]
    for name, side, total, color, hint in (sides if registries else []):
        table_sheet(wb, name, f"ИНН {inn} · {name.upper()} на {format_money(total)} · платежей", color, hint, HEAD, WIDTHS,
                    rows(side), money=(10,), dates=(5,))
    wb.save(path)
    res["path"] = path
    return res
