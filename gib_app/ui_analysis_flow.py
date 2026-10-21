"""Фильтр по файлам: файлы → у каждого своя карточка (лист, ИНН, сумма) → ИНН и/или порог суммы →
«Показать в Excel» → строки как есть из всех файлов. Анализ делает сам пользователь."""
import os

from PySide6.QtWidgets import QMessageBox

from file_sheets import read_files
from report_inn_search import clean_inn, filter_title, write_filter_report
from report_paths import filter_report_path
from ui_file_card import FileCard
from ui_result import show_result
from utils import format_money, to_number
from worker import Worker

ANALYSIS_PAGE = 3
TITLE = "Фильтр"


def start(win, paths: list):
    """Выбраны файлы: посмотреть их листы в фоне и показать по карточке на файл."""
    win.an_paths, win.an_mappings, win.an_file_cards = list(paths), {}, []
    _clear_cards(win)
    win.pages.setCurrentIndex(ANALYSIS_PAGE)
    _run(win, read_files, win.an_paths, done=lambda props: _show_cards(win, props))


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
    """Можно искать: хотя бы один файл взят и у всех взятых колонки указаны."""
    cards = getattr(win, "an_file_cards", [])
    return any(c.included() for c in cards) and not any(c.problem() for c in cards)


def _update_run(win):
    if not win.worker:
        win.an_run_btn.setEnabled(_ready(win))


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
    win.an_other_btn.setEnabled(True)
    _update_run(win)


def _threshold(win):
    """Порог суммы с экрана: (от, до) или None; неправильное число — сообщение и False."""
    texts = (win.an_min.text().strip(), win.an_max.text().strip())
    values = [to_number(t) if t else None for t in texts]
    if any(t and v is None for t, v in zip(texts, values)):
        QMessageBox.information(win, TITLE, "Впишите сумму числом, например 2000 или 2 000,50.")
        return False
    lo, hi = values
    if lo is not None and hi is not None and hi < lo:
        QMessageBox.information(win, TITLE, "«До» меньше, чем «от».")
        return False
    return (lo, hi) if lo is not None or hi is not None else None


def run(win):
    """«Показать в Excel»: строки всех взятых файлов по фильтру (ИНН и/или сумма) → отдельный Excel."""
    if win.worker or not _ready(win) or (threshold := _threshold(win)) is False:
        return
    inn = clean_inn(win.an_inn.text())
    if win.an_inn.text().strip() and not 8 <= len(inn) <= 14:
        QMessageBox.information(win, TITLE, "ИНН — цифрами, например 040058284.")
        return
    if not inn and not threshold:
        QMessageBox.information(win, TITLE, "Впишите ИНН и/или сумму (от … до …) — по ним покажем строки.")
        return
    win.an_mappings = {c.path: c.mapping() for c in win.an_file_cards if c.included()}
    _run(win, write_filter_report, filter_report_path(filter_title(inn, threshold)), list(win.an_mappings),
         win.an_mappings, inn, threshold, done=lambda res: _done(win, res))


def _done(win, res):
    if not res["path"]:
        QMessageBox.information(win, TITLE, f"По фильтру «{res['title']}» строк нет ни в одном файле.")
        return
    files = [f"{name}: <b>{count}</b> строк" + (f" на <b>{format_money(total)}</b>" if total is not None else "")
             for name, count, total in res["files"]]
    show_result(win, True, f"✓  {res['title']} — {res['rows']} строк", [
        f"Создан файл: <b>{os.path.basename(res['path'])}</b>",
        f"<span style='color:#64748B'>Папка: {os.path.dirname(res['path'])}</span>", ""] + files,
        path=res["path"], note="Исходные файлы не изменены.")
    win.back_page = ANALYSIS_PAGE


def back(win):
    if not win.worker:
        win.pages.setCurrentIndex(0)
