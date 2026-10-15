"""Фоновый рабочий поток QThread для длительных операций."""
from PySide6.QtCore import QThread, Signal

from excel_com import WriteError

# Запущенные потоки держим здесь до их полного завершения. Иначе, когда окно забывает поток сразу после
# сигнала done (а поток ещё не вышел из run), Python удаляет объект и Qt аварийно закрывает программу
# («QThread: Destroyed while thread is still running»). Так было при «Сводной»: два потока подряд.
_RUNNING = set()


class Worker(QThread):
    progress = Signal(int, str)
    done = Signal(object)            # результат задачи: словарь, список (просмотр файлов) и т.п.
    failed = Signal(str)

    def __init__(self, func, *args):
        super().__init__()
        self.func, self.args = func, args
        _RUNNING.add(self)
        self.finished.connect(lambda: _RUNNING.discard(self))

    def run(self):
        try:
            self.done.emit(self.func(*self.args, progress=lambda p, t: self.progress.emit(p, t)))
        except WriteError as e:
            self.failed.emit(str(e))
        except Exception as e:
            self.failed.emit(f"{type(e).__name__}: {e}")
