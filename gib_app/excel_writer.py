"""Заполнение кодов и сумм ГИБ в существующий файл через Excel COM."""
from constants import XL_CALC_MANUAL, XL_DOWN, XL_GREEN, XL_YELLOW
from excel_columns import col_values, ensure_gib_columns, gib_columns
from excel_com import WriteError, make_backup, open_workbook
from excel_reconcile import (
    fill_total_formulas, mark_no_gib, read_reconciliation,
    setup_reconciliation, verify_rows
)
from utils import to_number


def fill_in_place(path, sheet_name, purpose_col, header_row, items, first_row=None, last_row=None,
                  check_rows=(), no_gib_rows=(), progress=lambda p, t: None):
    """Суммы ставятся в те же строки, где стоят назначения."""
    import pythoncom
    import pywintypes
    import win32com.client

    if not items:
        raise WriteError("Нет строк для заполнения.")
    pythoncom.CoInitialize()
    progress(2, "Подключаюсь к Excel…")
    xl, wb, own = open_workbook(win32com.client, pywintypes, path)
    backup = None
    old_calc = old_screen = None
    try:
        if wb.ReadOnly:
            raise WriteError("Файл открыт только для чтения (возможно, его открыл другой пользователь).")
        progress(5, "Делаю копию файла…")
        backup = make_backup(path, wb, own)
        try:
            ws = wb.Worksheets(sheet_name)
        except pywintypes.com_error:
            raise WriteError(f"Лист «{sheet_name}» не найден.")

        old_screen = xl.ScreenUpdating
        xl.ScreenUpdating = False
        try:
            old_calc = xl.Calculation
            xl.Calculation = XL_CALC_MANUAL
        except pywintypes.com_error:
            pass

        header_inserted = False
        if header_row is None:
            first = min(r for r, _ in items)
            ws.Rows(first).Insert(Shift=XL_DOWN)
            header_row = first
            items = [(r + 1, g) for r, g in items]
            check_rows = [r + 1 for r in check_rows]
            no_gib_rows = [r + 1 for r in no_gib_rows]
            header_inserted = True
            first_row = first_row + 1 if first_row else None
            last_row = last_row + 1 if last_row else None
        first_row = min([first_row or 10 ** 9] + [r for r, _ in items])
        last_row = max([last_row or 0] + [r for r, _ in items])

        progress(10, "Добавляю сверку с выпиской…")
        recon = setup_reconciliation(ws, header_row, purpose_col, first_row, last_row)

        progress(15, "Проверяю столбцы ГИБ…")
        needed = sorted({code for _, gibs in items for code in gibs})
        created = ensure_gib_columns(ws, header_row, needed, purpose_col)

        progress(50, "Ставлю суммы…")
        gib_cols = gib_columns(ws, header_row)
        check = set(check_rows)
        mark_no_gib(ws, purpose_col, no_gib_rows, [r for r, _ in items])
        for r in check:
            ws.Cells(r, purpose_col).Interior.Color = XL_YELLOW
        r1 = min(r for r, _ in items)
        r2 = max(r for r, _ in items)
        replaced = 0
        for code, col in gib_cols.items():
            mine = {r: g[code] for r, g in items if code in g}
            if not mine:
                continue
            cur = col_values(ws, col, r1, r2, "Formula")
            out = list(cur)
            for r, amount in mine.items():
                old = cur[r - r1]
                if old not in ("", None) and to_number(old) != amount:
                    replaced += 1
                out[r - r1] = amount
            ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).Formula = [(v,) for v in out]
            for r in mine:
                ws.Cells(r, col).Interior.Color = XL_YELLOW if r in check else XL_GREEN
        for code in created:
            ws.Cells(header_row, gib_cols[code]).Interior.Color = XL_GREEN

        fill_total_formulas(ws, header_row, recon["total"], first_row, last_row)

        if old_calc is not None:
            xl.Calculation = old_calc
            old_calc = None
        xl.Calculate()
        progress(90, "Сверяю итоги…")
        check_res = verify_rows(ws, items, header_row, r2)
        check_res["recon"] = read_reconciliation(ws, header_row, first_row, last_row)
        progress(95, "Сохраняю…")
        wb.Save()
        progress(100, "Готово")
        return {
            "filled": len(items),
            "created": created,
            "replaced": replaced,
            "header_inserted": header_inserted,
            "backup": backup,
            "left_open": not own,
            "check": check_res,
        }
    finally:
        try:
            if old_calc is not None:
                xl.Calculation = old_calc
            if old_screen is not None:
                xl.ScreenUpdating = old_screen
        except Exception:
            pass
        if own:
            try:
                wb.Close(False)
            finally:
                xl.Quit()
