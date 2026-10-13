"""Управление сессиями Microsoft Excel COM и резервным копированием файлов."""
import os
import shutil
from datetime import datetime


class WriteError(Exception):
    pass


def make_backup(path: str, wb, own: bool) -> str:
    """Копия файла до изменений в папку «_резервные копии» рядом с ним."""
    folder = os.path.join(os.path.dirname(os.path.abspath(path)), "_резервные копии")
    name, ext = os.path.splitext(os.path.basename(path))
    dst = os.path.join(folder, f"{name} {datetime.now():%Y-%m-%d %H-%M-%S}{ext}")
    try:
        os.makedirs(folder, exist_ok=True)
        if own:
            shutil.copy2(path, dst)
        else:
            wb.SaveCopyAs(dst)
    except Exception as e:
        raise WriteError(f"Не удалось сделать копию файла, поэтому файл не изменён.\n{e}")
    return dst


def open_workbook(win32, pywintypes, path: str):
    """Если книга уже открыта у пользователя в Excel — пишем прямо в неё."""
    full = os.path.normcase(os.path.abspath(path))
    try:
        running = win32.GetActiveObject("Excel.Application")
        for wb in running.Workbooks:
            if os.path.normcase(wb.FullName) == full:
                return running, wb, False
    except (pywintypes.com_error, AttributeError):
        pass
    try:
        xl = win32.DispatchEx("Excel.Application")
    except pywintypes.com_error:
        raise WriteError("Не удалось запустить Microsoft Excel. Для записи нужен установленный Excel.")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0)
    except pywintypes.com_error as e:
        xl.Quit()
        raise WriteError(f"Excel не смог открыть файл: {e}")
    return xl, wb, True
