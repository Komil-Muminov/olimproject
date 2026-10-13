"""Управление сессиями Microsoft Excel COM и резервным копированием файлов."""
import os
import shutil
from datetime import datetime


class WriteError(Exception):
    pass


def backup_dir() -> str:
    """Папка «_резервные копии» в корне проекта: рядом с gib_app при запуске из исходников,
    рядом с GIB.exe (а если exe лежит в dist — уровнем выше) при запуске программы."""
    import sys
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.executable))
        if os.path.basename(base).lower() == "dist":
            base = os.path.dirname(base)
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "_резервные копии")


def make_backup(path: str, wb, own: bool) -> str:
    """Копия файла до изменений в папку «_резервные копии» в корне проекта."""
    folder = backup_dir()
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
