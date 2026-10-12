"""Собирает assets/gib_template.xlsm — пустую книгу с макросом распределения по ГИБ.

Запускается один раз разработчиком (нужен Excel с включённым «Доверять доступ к объектной
модели проектов VBA»). У пользователей программы эта настройка не нужна: программа берёт
готовый шаблон и копирует в него листы их файла.

    python assets/build_template.py <Макрос.xlsm>

Код макроса — assets/modGIB.bas.txt (модуль modGIB), обработчики событий книги берутся
из модуля ThisWorkbook исходного Макрос.xlsm.
"""
import os
import sys

import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))


def main(source):
    code = open(os.path.join(HERE, "modGIB.bas.txt"), encoding="utf-8").read().replace("\n", "\r\n")
    out = os.path.join(HERE, "gib_template.xlsm")
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    wb = xl.Workbooks.Open(os.path.abspath(source))
    try:
        keep = wb.Worksheets(wb.Worksheets.Count)
        for ws in list(wb.Worksheets):
            if ws.Name != keep.Name:
                ws.Delete()
        keep.Cells.Clear()
        for shp in list(keep.Shapes):
            shp.Delete()
        keep.Name = "_gib_tmp"
        for nm in list(wb.Names):
            nm.Delete()
        cm = wb.VBProject.VBComponents("modGIB").CodeModule
        cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(code)
        if os.path.exists(out):
            os.remove(out)
        wb.SaveAs(out, FileFormat=52)
        print("Готово:", out, "| строк в modGIB:", cm.CountOfLines)
    finally:
        wb.Close(False)
        xl.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
