"""Через Excel строит на листе «Сводная» настоящую сводную таблицу по листу «Данные» и срезы
(кнопки-фильтры). Библиотека openpyxl сводные создавать не умеет — это делает сам Excel."""
import os

from analysis_model import STATUSES
from excel_com import excel_app, set_number_format

XL_DATABASE, XL_ROW_FIELD, XL_COLUMN_FIELD, XL_SUM, XL_DESCENDING = 1, 1, 2, -4157, 2
PIVOT_ROW = 14                  # сводная — под срезами
TITLE = "СВОДНАЯ ТАБЛИЦА"
HINT = ("Нажимайте кнопки срезов — таблица сразу перестроится (Ctrl+клик — несколько). «+» слева раскрывает "
        "организации. Поля можно перетаскивать: правый клик по таблице → «Показать список полей».")
ROLE_HINT = ("В срезе «Кто» выбраны плательщики: если выбрать и получателей, каждый платёж посчитается дважды "
             "(у того, кто платил, и у того, кто получил).")


def build_pivot(path: str, n_rows: int, n_cols: int, with_sums: bool, with_role: bool, progress=lambda p, t: None):
    progress(45, "Excel строит сводную таблицу…")
    with excel_app() as xl:
        wb = xl.Workbooks.Open(os.path.abspath(path))
        try:
            data = wb.Worksheets("Данные")
            table = data.ListObjects.Add(1, data.Range(data.Cells(1, 1), data.Cells(n_rows, n_cols)), None, 1)
            table.Name = "Данные"
            sheet = wb.Worksheets.Add(Before=data)
            sheet.Name = "Сводная"
            sheet.Cells(1, 1).Value = TITLE
            sheet.Cells(1, 1).Font.Bold = True
            sheet.Cells(1, 1).Font.Size = 14
            sheet.Cells(2, 1).Value = HINT
            sheet.Cells(2, 1).Font.Italic = True
            if with_role:
                sheet.Cells(3, 1).Value = ROLE_HINT
                sheet.Cells(3, 1).Font.Color = 0x0B53B4          # тёмно-оранжевый (BGR)
            progress(60, "Excel строит сводную таблицу…")
            pt = wb.PivotCaches().Create(XL_DATABASE, "Данные", 6).CreatePivotTable(
                sheet.Range(f"A{PIVOT_ROW}"), "Сводная")
            _layout(pt, with_sums)
            sheet.Columns(1).ColumnWidth = 55       # до срезов: иначе они сдвинутся вместе со столбцом
            progress(80, "Добавляю срезы…")
            fields = (["Кто"] if with_role else []) + ["Что случилось", "Ошибки", "Часть"]
            for k, field in enumerate(fields):
                cache = wb.SlicerCaches.Add2(pt, field)
                if field == "Кто":
                    cache.SlicerItems("получатель").Selected = False      # иначе суммы вдвое
                slicer = cache.Slicers.Add(
                    SlicerDestination=sheet, Name=f"Срез {k + 1}", Caption=field)
                # позицию задаём после создания: параметры в Add Excel игнорирует и ставит срезы стопкой
                slicer.Top, slicer.Left, slicer.Width, slicer.Height = 48, 5 + k * 200, 190, 130
            sheet.Activate()
            progress(92, "Сохраняю…")
            wb.Save()
        finally:
            wb.Close(False)


def _layout(pt, with_sums: bool):
    """Строки: «Что случилось» (свёрнуто) → организации, крупные сверху; столбцы: части; значения: суммы."""
    status = pt.PivotFields("Что случилось")
    status.Orientation = XL_ROW_FIELD
    pt.PivotFields("Название").Orientation = XL_ROW_FIELD
    pt.PivotFields("Часть").Orientation = XL_COLUMN_FIELD
    value = pt.AddDataField(pt.PivotFields("Сумма" if with_sums else "Записей"),
                            "Сумма, всего" if with_sums else "Записей, всего", XL_SUM)
    set_number_format(value, "#,##0.00" if with_sums else "0")
    present = {str(item.Name) for item in status.PivotItems()}
    for pos, name in enumerate([s for s in STATUSES if s in present], 1):
        status.PivotItems(name).Position = pos            # порядок как в отчёте, а не по алфавиту
    pt.PivotFields("Название").AutoSort(XL_DESCENDING, value.Name)
    status.ShowDetail = False
    pt.CompactLayoutRowHeader = "Что случилось → организации"
    pt.CompactLayoutColumnHeader = "Части"
    pt.GrandTotalName = "Итого"
