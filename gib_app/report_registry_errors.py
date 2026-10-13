"""Листы ошибок отчёта «Свод по ИНН»: двойные платежи, неправильные ИНН, странные даты."""
import os
import re

from report_fast import table_sheet


INN_IN_TEXT = re.compile(r"(?<!\d)\d{9}(?!\d)")


def inn_in_name(name: str) -> str:
    """ИНН, вписанный в название («… ИНН 010010058»), — подсказка, каким он должен быть."""
    m = INN_IN_TEXT.search(name or "")
    return m.group(0) if m else ""


def _sheet(sheets, name):
    return next(x for x in sheets if x[0] == name)


def write_error_sheets(wb, reg, sheets):
    fname = lambda f: os.path.basename(reg.files[f].path)

    _, color, hint = _sheet(sheets, "Двойные платежи")
    rows = sorted(reg.dups, key=lambda d: -d[1].amount)
    table_sheet(wb, "Двойные платежи", "ДВОЙНЫЕ ПЛАТЕЖИ", color,
                f"Платёж повторяется: {hint}. Крупные сверху. В последних столбцах — где он был в первый раз.",
                ["№", "Файл", "Строка", "№ док.", "Дата", "Плательщик", "ИНН плат.", "Получатель", "ИНН получ.",
                 "Сумма", "Первый раз: файл", "строка"],
                [7, 14, 9, 13, 12, 36, 12, 36, 12, 16, 14, 9],
                [[i, fname(f), p.row, p.doc, p.date, p.payer, p.payer_inn, p.rec, p.rec_inn, p.amount,
                  fname(first[0]), first[1]] for i, (f, p, first) in enumerate(rows, 1)],
                money=(10,), dates=(5,))

    _, color, hint = _sheet(sheets, "Неправильные ИНН")
    rows = sorted(reg.bad_inn, key=lambda b: -b[1].amount)
    table_sheet(wb, "Неправильные ИНН", "НЕПРАВИЛЬНЫЕ ИНН", color,
                f"{hint}. Крупные платежи сверху.",
                ["№", "Файл", "Строка", "Чей ИНН", "Что стоит", "Название", "ИНН из названия", "№ док.", "Дата",
                 "Сумма"],
                [7, 14, 9, 13, 14, 46, 14, 13, 12, 16],
                [[i, fname(f), p.row, side, inn or "(пусто)", name, inn_in_name(name), p.doc, p.date, p.amount]
                 for i, (f, p, side, inn) in enumerate(rows, 1)
                 for name in [p.payer if side == "плательщик" else p.rec]],
                money=(10,), dates=(9,))

    _, color, hint = _sheet(sheets, "Странные даты")
    rows = sorted(reg.odd_dates, key=lambda x: (x[1].date, -x[1].amount))
    table_sheet(wb, "Странные даты", "СТРАННЫЕ ДАТЫ", color,
                f"{hint}. Возможно, опечатка в дате платёжки.",
                ["№", "Файл", "Период файла", "Строка", "№ док.", "Дата документа", "Плательщик", "Получатель", "Сумма"],
                [7, 14, 15, 9, 13, 14, 40, 40, 16],
                [[i, fname(f), reg.files[f].label, p.row, p.doc, p.date, p.payer, p.rec, p.amount]
                 for i, (f, p) in enumerate(rows, 1)],
                money=(9,), dates=(6,))
