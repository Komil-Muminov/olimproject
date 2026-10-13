"""Собирает assets/gib_template.xlsm — пустую книгу с макросом распределения по ГИБ.

Запускается один раз разработчиком (нужен Excel с включённым «Доверять доступ к объектной
модели проектов VBA»). У пользователей программы эта настройка не нужна: программа берёт
готовый шаблон и копирует в него листы их файла.

    python assets/build_template.py <любая книга .xlsm с модулем modGIB>

Код макроса — assets/modGIB.bas.txt (модуль modGIB) и assets/ThisWorkbook.cls.txt
(обработчики событий книги). Оба взяты из «Банк+муҳосибӣ (Бучет) -ислох шуда.xlsm»
с нашими доработками: помеченный лист (GIB_ON/GIB_HDR/GIB_COL), «Основание»,
конец данных на «Обороты за период», жёлтая подсветка кода из 5+ цифр.
"""
import os
import sys

import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))


def _read(name):
    return open(os.path.join(HERE, name), encoding="utf-8").read().replace("\n", "\r\n")


def main(source):
    out = os.path.join(HERE, "gib_template.xlsm")
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    wb = xl.Workbooks.Open(os.path.abspath(source))
    try:
        keep = next(ws for ws in wb.Worksheets if ws.Visible == -1)
        for ws in list(wb.Worksheets):
            if ws.Name != keep.Name:
                ws.Visible = -1
                ws.Delete()
        keep.Cells.Clear()
        for shp in list(keep.Shapes):
            shp.Delete()
        keep.Name = "_gib_tmp"
        for nm in list(wb.Names):
            try:
                nm.Delete()
            except Exception:
                pass        # служебные имена Excel (_xlfn…) не удаляются — не мешают
        for name, text in (("modGIB", _read("modGIB.bas.txt")), ("ThisWorkbook", _read("ThisWorkbook.cls.txt"))):
            cm = wb.VBProject.VBComponents(name).CodeModule
            cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(text)
        lines = wb.VBProject.VBComponents("modGIB").CodeModule.CountOfLines
        if os.path.exists(out):
            os.remove(out)
        wb.SaveAs(out, FileFormat=52)
        print("Готово:", out, "| строк в modGIB:", lines)
    finally:
        wb.Close(False)
        xl.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
