"""Журнал ошибок программы: «ошибки программы.log» в корне проекта (рядом с GIB.exe).

Записывается любая необработанная ошибка (с подробностями) и даже аварийное падение — чтобы можно было
понять причину, если программа закрылась сама."""
import faulthandler
import os
import sys
import traceback
from datetime import datetime

from app_paths import project_dir

LOG_NAME = "ошибки программы.log"
_log_file = None


def log_path() -> str:
    return os.path.join(project_dir(), LOG_NAME)


def install(show_error=None):
    """Включить журнал. show_error(текст) — показать ошибку пользователю (например, окном)."""
    global _log_file
    try:
        _log_file = open(log_path(), "a", encoding="utf-8")
        faulthandler.enable(file=_log_file)          # аварийное падение (без Python-ошибки) — тоже в журнал
    except OSError:
        _log_file = None

    def hook(kind, value, tb):
        text = "".join(traceback.format_exception(kind, value, tb))
        if _log_file:
            _log_file.write(f"\n=== {datetime.now():%d.%m.%Y %H:%M:%S} ===\n{text}")
            _log_file.flush()
        if show_error:
            show_error(f"{kind.__name__}: {value}\n\nПодробности записаны в «{LOG_NAME}».")

    sys.excepthook = hook
