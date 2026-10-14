"""Управление сессиями Microsoft Excel COM и резервным копированием файлов."""
import os
import shutil
from contextlib import contextmanager
from datetime import datetime

from app_paths import backup_dir


class WriteError(Exception):
    pass


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


@contextmanager
def excel_workbook(path: str):
    """Книга для записи: (xl, wb, own). Макросы книги отключены на время работы; в конце
    настройки Excel возвращаются, а своя (не открытая пользователем) книга и Excel закрываются."""
    import pythoncom
    import pywintypes
    import win32com.client

    pythoncom.CoInitialize()
    xl, wb, own = open_workbook(win32com.client, pywintypes, path)
    saved = {}
    try:
        saved = {k: getattr(xl, k) for k in ("EnableEvents", "ScreenUpdating", "Calculation")}
    except Exception:
        pass
    try:
        if wb.ReadOnly:
            raise WriteError("Файл открыт только для чтения (возможно, его открыл другой пользователь).")
        xl.EnableEvents = False
        yield xl, wb, own
    finally:
        for k, v in saved.items():
            try:
                setattr(xl, k, v)
            except Exception:
                pass
        if own:
            try:
                wb.Close(False)
            finally:
                xl.Quit()


@contextmanager
def excel_app():
    """Отдельный скрытый Excel; в конце закрывается (книги закрывает вызывающий код)."""
    import pythoncom
    import pywintypes
    import win32com.client

    pythoncom.CoInitialize()
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
    except pywintypes.com_error:
        raise WriteError("Не удалось запустить Microsoft Excel.")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.ScreenUpdating = False
    try:
        yield xl
    finally:
        xl.Quit()


def get_sheet(wb, name: str):
    try:
        return wb.Worksheets(name)
    except Exception:
        raise WriteError(f"Лист «{name}» не найден.")


EN_LCID = 0x409


def _invoke_fmt(rng, put: bool, fmt=None):
    """Формат чисел с английским языком, как его видит VBA («0.00», «General»). Обычное
    rng.NumberFormat через pywin32 работает с языком Windows: «0.00» отклоняется, «General» = «Основной»."""
    import pythoncom
    obj = rng._oleobj_
    disp = obj.GetIDsOfNames("NumberFormat")
    if put:
        return obj.Invoke(disp, EN_LCID, pythoncom.DISPATCH_PROPERTYPUT, 0, fmt)
    return obj.Invoke(disp, EN_LCID, pythoncom.DISPATCH_PROPERTYGET, 1)


def number_format(rng):
    """Формат диапазона (None — в диапазоне разные форматы)."""
    return _invoke_fmt(rng, False)


def set_number_format(rng, fmt: str):
    _invoke_fmt(rng, True, fmt)
