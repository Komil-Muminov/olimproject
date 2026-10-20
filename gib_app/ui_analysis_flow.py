"""Глубокий анализ: файлы → у каждого своя карточка с колонками → «Анализ» → чтение в фоне → Excel → итог."""
import os

from PySide6.QtWidgets import QDialog, QMessageBox

from analysis_jobs import make_report
from analysis_propose import propose_files
from analysis_scan import auto_mode, scan_files
from report_inn_search import clean_inn, write_inn_report
from report_paths import analysis_report_path, inn_report_path
from ui_analysis_result import result_view
from ui_file_card import FileCard
from ui_pair_dialog import PairDialog, needs_pair
from ui_result import show_result
from utils import format_money
from worker import Worker

ANALYSIS_PAGE = 3


def start(win, paths: list):
    """Выбраны файлы: посмотреть их листы в фоне и показать по карточке на файл."""
    win.an_paths, win.an_mappings, win.an_file_cards = list(paths), {}, []
    _clear_cards(win)
    win.pages.setCurrentIndex(ANALYSIS_PAGE)
    _run(win, propose_files, win.an_paths, done=lambda props: _show_cards(win, props))


def _clear_cards(win):
    while win.an_cards.count() > 1:                 # последний элемент — растяжка
        item = win.an_cards.takeAt(0)
        if item.widget():
            item.widget().deleteLater()


def _show_cards(win, props: list):
    for prop in props:
        card = FileCard(prop)
        card.changed.connect(lambda: _update_run(win))
        win.an_cards.insertWidget(win.an_cards.count() - 1, card)
        win.an_file_cards.append(card)
    _update_run(win)


def _ready(win) -> bool:
    """«Анализ» можно: хотя бы один файл взят и у всех взятых колонки указаны."""
    cards = getattr(win, "an_file_cards", [])
    return any(c.included() for c in cards) and not any(c.problem() for c in cards)


def _update_run(win):
    if not win.worker:
        for b in (win.an_run_btn, win.an_inn_btn):
            b.setEnabled(_ready(win))


def _run(win, func, *args, done):
    for b in (win.an_run_btn, win.an_other_btn, win.an_inn_btn):
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
    win.an_other_btn.setEnabled(True)
    _update_run(win)


def run(win):
    """«Анализ»: колонки из карточек (только взятые файлы) → чтение в фоне → Excel."""
    if win.worker or not _ready(win):
        return
    win.an_mappings = _chosen(win)
    _run(win, scan_files, list(win.an_mappings), win.an_mappings, done=lambda res: _scanned(win, res))


def _chosen(win) -> dict:
    """Колонки из карточек — только взятые файлы."""
    return {c.path: c.mapping() for c in win.an_file_cards if c.included()}


def search_inn(win):
    """«Один ИНН во всех файлах»: все его строки из всех взятых файлов → отдельный Excel."""
    inn = clean_inn(win.an_inn.text())
    if win.worker or not _ready(win):
        return
    if not 8 <= len(inn) <= 14:
        QMessageBox.information(win, "Один ИНН во всех файлах", "Введите ИНН цифрами, например 040058284.")
        return
    win.an_mappings = _chosen(win)
    _run(win, write_inn_report, inn_report_path(inn), list(win.an_mappings), win.an_mappings, inn,
         done=lambda res: _inn_done(win, res))


def _inn_done(win, res):
    if not res["path"]:
        QMessageBox.information(win, "Один ИНН во всех файлах", f"ИНН {res['inn']} не найден ни в одном файле.")
        return
    files = [f"{name}: <b>{count}</b> строк" + (f" на <b>{format_money(total)}</b>" if total is not None else "")
             for name, count, total in res["files"]]
    show_result(win, True, f"✓  ИНН {res['inn']} — {res['rows']} строк", [
        f"Создан файл: <b>{os.path.basename(res['path'])}</b>",
        f"<span style='color:#64748B'>Папка: {os.path.dirname(res['path'])}</span>", ""] + files,
        path=res["path"], note="Исходные файлы не изменены.")
    win.back_page = ANALYSIS_PAGE


def _scanned(win, res):
    raw = res["raw"]
    mode, pair = auto_mode(raw), None       # по умолчанию: первая часть → последняя
    if needs_pair(raw, mode):
        dlg = PairDialog(win, raw, mode)
        if dlg.exec() != QDialog.Accepted:
            return
        pair = dlg.pair()
    _run(win, make_report, analysis_report_path(), raw, mode, pair, win.an_mappings,
         done=lambda r: _done(win, r, r["data"]))


def _done(win, res, data):
    lines, chips, errors = result_view(res, data)
    show_result(win, True, "✓  Анализ готов", lines, path=res["path"], chips=chips, after=errors)
    win.back_page = ANALYSIS_PAGE


def back(win):
    if not win.worker:
        win.pages.setCurrentIndex(0)
