"""Фоновый рабочий поток QThread для длительных операций."""
from PySide6.QtCore import QThread, Signal

from excel_com import WriteError


class Worker(QThread):
    progress = Signal(int, str)
    done = Signal(dict)
    failed = Signal(str)

    def __init__(self, func, *args):
        super().__init__()
        self.func, self.args = func, args

    def run(self):
        try:
            self.done.emit(self.func(*self.args, progress=lambda p, t: self.progress.emit(p, t)))
        except WriteError as e:
            self.failed.emit(str(e))
        except Exception as e:
            self.failed.emit(f"{type(e).__name__}: {e}")
