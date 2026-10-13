"""Встраивание макроса автоматики и элементов управления в книгу Excel."""
import os
from constants import XL_DOWN, XL_XLSM
from excel_com import WriteError, excel_app
from excel_reconcile import setup_reconciliation


# кнопки — как в «Банк+муҳосибӣ (Бучет) -ислох шуда.xlsm»: (имя, подпись, макрос, x, ширина)
BUTTONS = (
    ("btnGibRecalc", "Пересчитать ГИБ", "GIB_RecalcSheet", 3, 118),
    ("btnGibAddRows", "Добавить строки", "GIB_AddRows", 125, 124),
    ("btnGibFreeze", "Зафиксировать распределение", "GIB_FreezeDistribution", 253, 180),
    ("btnGibUnfreeze", "Снять фиксацию", "GIB_UnfreezeDistribution", 437, 120),
)


def add_controls(ws):
    """Кнопки макроса в верхней строке и красная надпись «Макросы выключены»."""
    for name, caption, macro, x, width in BUTTONS:
        btn = ws.Buttons().Add(x, 4, width, 22)
        btn.Name = name
        btn.Caption = caption
        btn.OnAction = macro
        btn.PrintObject = False

    warn = ws.Shapes.AddTextbox(1, 2, 1, 640, 28)
    warn.Name = "gibMacroWarning"
    warn.Fill.ForeColor.RGB = 38 * 256 * 256 + 38 * 256 + 220
    warn.Line.Visible = False
    tf = warn.TextFrame2
    tf.TextRange.Text = ("⚠ Макросы выключены — нажмите «Включить содержимое» на жёлтой панели вверху. "
                         "Без этого ГИБ не заполняются.")
    tf.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
    tf.TextRange.Font.Bold = True
    tf.TextRange.Font.Size = 10
    tf.VerticalAnchor = 3
    try:
        warn.DrawingObject.PrintObject = False
    except Exception:
        pass


def embed_automation(src, dst, template, sheet_name, header_row, purpose_col, progress=lambda p, t: None):
    """«Встроить автоматику в Excel»: копия файла в формате .xlsm с макросом распределения по ГИБ."""
    if os.path.normcase(os.path.abspath(src)) == os.path.normcase(os.path.abspath(dst)):
        raise WriteError("Новый файл должен отличаться от исходного.")
    if not os.path.exists(template):
        raise WriteError(f"Не найден шаблон с макросом:\n{template}")
    progress(5, "Запускаю Excel…")
    with excel_app() as xl:
        report = _embed(xl, src, dst, template, sheet_name, header_row, purpose_col, progress)
    progress(100, "Готово")
    return {"path": os.path.abspath(dst), "report": report}


def _embed(xl, src, dst, template, sheet_name, header_row, purpose_col, progress) -> str:
    """Копирует листы в шаблон с макросом, сохраняет .xlsm и запускает распределение."""
    book = source = None
    try:
        book = xl.Workbooks.Open(os.path.abspath(template), ReadOnly=True)
        source = xl.Workbooks.Open(os.path.abspath(src), ReadOnly=True, UpdateLinks=0)
        progress(20, "Копирую листы…")
        source.Sheets.Copy(Before=book.Sheets(1))
        book.Worksheets("_gib_tmp").Delete()
        ws = book.Worksheets(sheet_name)

        if xl.WorksheetFunction.CountA(ws.Rows(1)) > 0 or ws.Range("A1").MergeCells:
            ws.Rows(1).Insert(Shift=XL_DOWN)
            header_row += 1
        if ws.Rows(1).RowHeight < 30:
            ws.Rows(1).RowHeight = 30

        setup_reconciliation(ws, header_row, purpose_col, header_row + 1, None)

        for name, value in (("GIB_ON", 1), ("GIB_HDR", header_row), ("GIB_COL", purpose_col)):
            ws.Names.Add(Name=name, RefersTo=f"={value}", Visible=False)

        progress(35, "Добавляю кнопку…")
        add_controls(ws)

        source.Close(False)
        source = None
        progress(50, "Сохраняю…")
        if os.path.exists(dst):
            os.remove(dst)
        book.SaveAs(os.path.abspath(dst), FileFormat=XL_XLSM)
        for link in book.LinkSources(1) or ():
            if os.path.normcase(link) == os.path.normcase(os.path.abspath(src)):
                book.ChangeLink(link, os.path.abspath(dst), 1)

        progress(65, "Заполняю ГИБ…")
        xl.EnableEvents = True
        book.Activate()
        ws.Activate()
        xl.Run("GIB_RecalcSheet")
        report = xl.Run("GIB_LastReport")
        if str(report).startswith("ERROR"):
            raise WriteError(f"Макрос сообщил об ошибке: {report[7:]}")
        progress(90, "Сохраняю…")
        book.Save()
        return report
    finally:
        if source is not None:
            source.Close(False)
        if book is not None:
            book.Close(False)
