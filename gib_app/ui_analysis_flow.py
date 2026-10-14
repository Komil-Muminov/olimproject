"""Глубокий анализ: файлы → «Анализировать» → чтение в фоне → Excel → итог простыми словами."""
import os

from PySide6.QtWidgets import QDialog, QMessageBox

from analysis_model import FLAGS, STATUSES
from analysis_scan import AS_LIST, scan_files
from report_analysis import write_analysis_report
from report_paths import analysis_report_path
from ui_column_dialog import ColumnDialog
from ui_result import show_result
from worker import Worker

ANALYSIS_PAGE = 3


def _int(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def start(win, paths: list):
    """Выбраны файлы: показать их и кнопку «Анализировать»."""
    win.an_paths, win.an_mappings = list(paths), {}
    win.an_files.setText("📚  " + ",  ".join(os.path.basename(p) for p in paths))
    win.pages.setCurrentIndex(ANALYSIS_PAGE)


def _run(win, func, *args, done):
    win.an_run_btn.setEnabled(False)
    win.an_other_btn.setEnabled(False)
    win.an_progress.setVisible(True)
    win.an_progress.setValue(0)
    win.worker = Worker(func, *args)
    win.worker.progress.connect(lambda v, t: (win.an_progress.setValue(v), win.an_progress.setFormat(f"{t}  %p%")))
    win.worker.done.connect(lambda res: (_idle(win), done(res)))
    win.worker.failed.connect(lambda err: (_idle(win), win._failed(err)))
    win.worker.start()


def _idle(win):
    win.worker = None
    win.an_progress.setVisible(False)
    win.an_run_btn.setEnabled(True)
    win.an_other_btn.setEnabled(True)


def run(win):
    if not win.worker:
        _run(win, scan_files, win.an_paths, win.an_mappings, done=lambda res: _scanned(win, res))


def _scanned(win, res):
    if "needs_mapping" in res:
        return _ask_columns(win, res["needs_mapping"], res["guess"])
    _run(win, write_analysis_report, analysis_report_path(), res["data"], done=lambda r: _done(win, r, res["data"]))


def _ask_columns(win, path: str, guess: dict):
    q = QMessageBox(QMessageBox.Question, "Как читать файл?", f"«{os.path.basename(path)}» похож на реестр "
                    "платежей, но программа узнала не все столбцы.", parent=win)
    manual = q.addButton("Указать столбцы", QMessageBox.AcceptRole)
    as_list = q.addButton("Читать как список организаций", QMessageBox.ActionRole)
    q.addButton("Отмена", QMessageBox.RejectRole)
    q.exec()
    if q.clickedButton() is manual:
        dlg = ColumnDialog(win, path, guess)
        if dlg.exec() != QDialog.Accepted:
            return
        win.an_mappings[path] = dlg.mapping()
    elif q.clickedButton() is as_list:
        win.an_mappings[path] = AS_LIST
    else:
        return
    run(win)


def _done(win, res, data):
    st, fl = res["statuses"], res["flags"]
    parts = " → ".join(p.label for p in data.parts)
    lines = [f"Создан файл: <b>{os.path.basename(res['path'])}</b>",
             f"<span style='color:#64748B'>Папка: {os.path.dirname(res['path'])}</span>", "",
             f"Организаций: <b>{_int(res['rows'])}</b> · частей: {len(data.parts)} ({parts})"]
    if st:
        lines.append("Что случилось: " + " · ".join(f"{s} <b>{_int(st[s])}</b>" for s in STATUSES if st[s]))
    lines.append("Ошибки: " + (" · ".join(f"{f} <b>{_int(fl[f])}</b>" for f in FLAGS if fl[f]) or "не найдено ✓"))
    show_result(win, True, "✓  Анализ готов", lines, path=res["path"], note="Исходные файлы не изменены.")
    win.back_page = ANALYSIS_PAGE


def back(win):
    if not win.worker:
        win.pages.setCurrentIndex(0)
