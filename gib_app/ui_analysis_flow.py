"""Глубокий анализ: файлы → одно окно «Листы и ИНН» → чтение в фоне → Excel → итог простыми словами."""
import os

from PySide6.QtWidgets import QDialog, QMessageBox

from analysis_jobs import make_report
from analysis_propose import propose_files
from analysis_scan import auto_mode, scan_files
from report_paths import analysis_report_path
from ui_analysis_result import result_view
from ui_result import show_result
from ui_sheets_dialog import SheetsDialog
from worker import Worker

ANALYSIS_PAGE = 3


def start(win, paths: list):
    """Выбраны файлы: показать их и кнопку «Анализировать»."""
    win.an_paths, win.an_mappings = list(paths), {}
    win.an_files.setText("📚  " + ",  ".join(os.path.basename(p) for p in paths))
    win.pages.setCurrentIndex(ANALYSIS_PAGE)


def _run(win, func, *args, done):
    for b in (win.an_run_btn, win.an_other_btn):
        b.setEnabled(False)
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
    for b in (win.an_run_btn, win.an_other_btn):
        b.setEnabled(True)


def run(win):
    """«Анализ»: посмотреть файлы → окно «Листы и ИНН» (решает пользователь) → Excel."""
    if not win.worker:
        _run(win, propose_files, win.an_paths, done=lambda props: _ask_sheets(win, props, lambda used: _run(
            win, scan_files, used, win.an_mappings, done=lambda res: _scanned(win, res))))


def _ask_sheets(win, props: list, after):
    """Одно окно на все листы всех файлов; пустые файлы — мимо (сказать); выбрано — after(пути)."""
    empty = [os.path.basename(p["path"]) for p in props if not p["sheets"]]
    props = [p for p in props if p["sheets"]]
    if not props:
        QMessageBox.warning(win, "Анализ", "Все файлы пустые:\n" + "\n".join(empty))
        return
    if empty:
        QMessageBox.information(win, "Анализ", "Пустые файлы пропущены:\n" + "\n".join(empty))
    dlg = SheetsDialog(win, props)
    if dlg.exec() != QDialog.Accepted:
        return
    win.an_mappings = dlg.mappings()
    after([p["path"] for p in props if p["path"] in win.an_mappings])


def _scanned(win, res):
    raw = res["raw"]
    _run(win, make_report, analysis_report_path(), raw, auto_mode(raw), win.an_mappings,
         done=lambda r: _done(win, r, r["data"]))


def _done(win, res, data):
    lines, chips, errors = result_view(res, data)
    show_result(win, True, "✓  Анализ готов", lines, path=res["path"], chips=chips, after=errors)
    win.back_page = ANALYSIS_PAGE


def back(win):
    if not win.worker:
        win.pages.setCurrentIndex(0)
