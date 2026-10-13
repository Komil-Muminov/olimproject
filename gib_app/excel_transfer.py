"""«Заполнить ГИБ» из двух файлов: платежи выписки дописываются в таблицу второго файла
(после последней заполненной строки, до «Чамъ»), затем суммы раскладываются по ГИБ."""
from datetime import datetime

from excel_com import WriteError, make_backup, open_workbook
from excel_writer import fill_in_place
from statement import doc_text, read_statement
from target_layout import create_table, existing_keys, locate, make_room, payment_key


def _put(ws, col, r1, values, text=False):
    if col and values:
        rng = ws.Range(ws.Cells(r1, col), ws.Cells(r1 + len(values) - 1, col))
        if text:
            try:
                rng.NumberFormat = "@"      # текстом, как в выписке (свой формат даты ломает некоторые книги)
            except Exception:
                pass
        rng.Value = [(v,) for v in values]


def _date(v):
    return v.strftime("%d.%m.%Y") if isinstance(v, datetime) else v


def _write_rows(ws, t, start, rows):
    _put(ws, t.purpose, start, [p.purpose for p in rows])
    _put(ws, t.rash, start, [p.amount for p in rows])
    # № документа — текстом, как в выписке (с нулями впереди)
    _put(ws, t.doc, start, [doc_text(p.doc) if isinstance(p.doc, float) else p.doc for p in rows], True)
    _put(ws, t.date, start, [_date(p.date) for p in rows], True)
    if t.num:       # нумерация р/т по порядку от начала таблицы
        _put(ws, t.num, start, [start + i - t.first_row + 1 for i in range(len(rows))])


def transfer_and_fill(statement_path, target_path, sheet_name, header_row, purpose_col,
                      progress=lambda p, t: None):
    import pythoncom
    import pywintypes
    import win32com.client

    progress(2, "Читаю выписку…")
    _, payments = read_statement(statement_path)
    if not payments:
        raise WriteError("В выписке не нашлось платежей (нужны столбцы «Основание» и «Расходы»).")
    pythoncom.CoInitialize()
    progress(5, "Открываю файл, куда заполнять…")
    xl, wb, own = open_workbook(win32com.client, pywintypes, target_path)
    old_events = xl.EnableEvents
    try:
        if wb.ReadOnly:
            raise WriteError("Файл открыт только для чтения (возможно, его открыл другой пользователь).")
        xl.EnableEvents = False
        progress(8, "Делаю копию файла…")
        backup = make_backup(target_path, wb, own)
        try:
            ws = wb.Worksheets(sheet_name)
        except pywintypes.com_error:
            raise WriteError(f"Лист «{sheet_name}» не найден.")
        if header_row is None:          # пустой лист — сначала создаём таблицу
            header_row, purpose_col = create_table(ws)
        t = locate(ws, header_row, purpose_col)
        use_doc = bool(t.doc) and any(p.doc not in (None, "") for p in payments)
        seen = existing_keys(ws, t, use_doc)
        new, skipped = [], 0
        for p in payments:
            k = payment_key(p.purpose, p.amount, p.doc if use_doc else None)
            if seen[k] > 0:
                seen[k] -= 1
                skipped += 1
            else:
                new.append(p)
        progress(15, "Переношу платежи…")
        old_sum = t.sum_row
        start = make_room(ws, t, len(new)) if new else t.last_filled + 1
        if new:
            _write_rows(ws, t, start, new)
        wb.Save()
    finally:
        try:
            xl.EnableEvents = old_events
        except Exception:
            pass
        if own:
            try:
                wb.Close(False)
            finally:
                xl.Quit()

    res = {"added": len(new), "skipped": skipped, "total": len(payments), "backup": backup,
           "rows": (start, start + len(new) - 1) if new else None,
           "inserted": (t.sum_row - old_sum) if old_sum else 0, "sheet": sheet_name,
           "no_gib": [start + i for i, p in enumerate(new) if not p.gibs],
           "long_codes": [(start + i, p.longs) for i, p in enumerate(new) if p.longs], "fill": None}
    items = [(start + i, p.gibs) for i, p in enumerate(new) if p.gibs]
    if items:
        last = (t.sum_row - 1) if t.sum_row else start + len(new) - 1
        res["fill"] = fill_in_place(target_path, sheet_name, purpose_col, header_row, items, t.first_row, last,
                                    [r for r, _ in res["long_codes"]], res["no_gib"],
                                    progress=lambda v, s: progress(int(20 + v * 0.8), s), backup=False)
    progress(100, "Готово")
    return res
