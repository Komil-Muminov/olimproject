"""Заполнение кодов и сумм ГИБ в существующий файл через Excel COM."""
from constants import XL_CALC_MANUAL, XL_DOWN, XL_GREEN, XL_YELLOW
from excel_columns import col_values, ensure_gib_columns, gib_columns
from excel_com import WriteError, make_backup, open_workbook
from excel_reconcile import (
    fill_total_formulas, read_reconciliation, setup_reconciliation, update_balance_notes
)
from excel_reader import cell_name
from excel_verify import mark_no_gib, verify_rows
from utils import to_number


def fill_in_place(path, sheet_name, purpose_col, header_row, items, first_row=None, last_row=None,
                  check_rows=(), no_gib_rows=(), progress=lambda p, t: None, backup=True):
    """Суммы ставятся в те же строки, где стоят назначения. backup=False — копию уже сделали."""
    import pythoncom
    import pywintypes
    import win32com.client

    if not items:
        raise WriteError("Нет строк для заполнения.")
    pythoncom.CoInitialize()
    progress(2, "Подключаюсь к Excel…")
    xl, wb, own = open_workbook(win32com.client, pywintypes, path)
    make_copy, backup = backup, None
    old_calc = old_screen = None
    old_events = xl.EnableEvents
    try:
        if wb.ReadOnly:
            raise WriteError("Файл открыт только для чтения (возможно, его открыл другой пользователь).")
        xl.EnableEvents = False             # макросы книги (если есть) не мешают записи
        if make_copy:
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
        needed = sorted({code for _, gibs in items for code in gibs}, key=int)
        created = ensure_gib_columns(ws, header_row, needed, purpose_col, start_col=recon["rash"] + 1)

        progress(50, "Ставлю суммы…")
        gib_cols = gib_columns(ws, header_row)
        check = set(check_rows)
        mark_no_gib(ws, purpose_col, no_gib_rows, [r for r, _ in items])
        for r in check:
            ws.Cells(r, purpose_col).Interior.Color = XL_YELLOW
        r1 = min(r for r, _ in items)
        r2 = max(r for r, _ in items)
        replaced = 0
        item_rows = {r: g for r, g in items}
        for code, col in gib_cols.items():
            mine = {r: g[code] for r, g in items if code in g}
            cur = col_values(ws, col, r1, r2, "Formula")
            out = list(cur)
            # как в образце: строка заполняется заново по тексту — суммы ГИБ, которых
            # в назначении больше нет, стираются (формулы не трогаем)
            cleared = []
            for r, g in item_rows.items():
                old = cur[r - r1]
                if code not in g and old not in ("", None) and not str(old).startswith("="):
                    out[r - r1] = ""
                    cleared.append(r)
                    replaced += 1
            if not mine and out == list(cur):
                continue
            for r, amount in mine.items():
                old = cur[r - r1]
                if old not in ("", None) and to_number(old) != amount:
                    replaced += 1
                out[r - r1] = amount
            ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).Formula = [(v,) for v in out]
            for r in mine:
                ws.Cells(r, col).Interior.Color = XL_YELLOW if r in check else XL_GREEN
            for r in cleared:
                if int(ws.Cells(r, col).Interior.Color) in (XL_GREEN, XL_YELLOW):
                    ws.Cells(r, col).Interior.Pattern = -4142      # наша подсветка больше не нужна
        for code in created:
            ws.Cells(header_row, gib_cols[code]).Interior.Color = XL_GREEN
        if recon["sum_row"]:        # итог «Чамъ» у столбцов ГИБ, где его ещё нет
            for col in gib_cols.values():
                cell = ws.Cells(recon["sum_row"], col)
                if cell.Formula in ("", None):
                    cell.Formula = f"=SUM({cell_name(first_row, col)}:{cell_name(recon['sum_row'] - 1, col)})"

        fill_total_formulas(ws, header_row, recon["total"], first_row, last_row, recon["rash"])

        if old_calc is not None:
            xl.Calculation = old_calc
            old_calc = None
        xl.Calculate()
        # пояснение к балансу «Расходы»: откуда разница (видно при наведении)
        update_balance_notes(ws, recon["total"], recon["rash"], first_row, recon["sum_row"] - 1 if recon["sum_row"] else last_row)
        progress(90, "Сверяю итоги…")
        check_res = verify_rows(ws, items, header_row, r2)
        check_res["recon"] = read_reconciliation(ws, header_row, first_row, last_row, recon["rash"])
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
            xl.EnableEvents = old_events
        except Exception:
            pass
        if own:
            try:
                wb.Close(False)
            finally:
                xl.Quit()
