"""Глубокий анализ: файлы → для каждого «Как читать файл» → чтение в фоне → Excel → итог простыми словами."""
import os

from PySide6.QtWidgets import QDialog, QMessageBox

from analysis_jobs import make_report
from analysis_propose import propose_files
from analysis_scan import auto_mode, scan_files
from report_inn import clean_inn, write_inn_payments
from report_paths import analysis_report_path, inn_report_path
from ui_analysis_result import result_view
from ui_file_dialog import FileDialog
from ui_result import show_result
from utils import format_money
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
    for b in (win.an_run_btn, win.an_other_btn, win.an_pay_btn):
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
    for b in (win.an_run_btn, win.an_other_btn, win.an_pay_btn):
        b.setEnabled(True)


def run(win):
    """«Анализ»: посмотреть файлы → для каждого окно «Как читать файл» (решает пользователь) → Excel."""
    if not win.worker:
        _run(win, propose_files, win.an_paths, done=lambda props: _ask_files(win, props))


def _ask_files(win, props: list):
    mappings = {}
    for k, prop in enumerate(props, 1):
        if not prop["sheets"]:
            QMessageBox.warning(win, "Анализ", f"Файл «{os.path.basename(prop['path'])}» пустой.")
            return
        dlg = FileDialog(win, prop, k, len(props))
        if dlg.exec() != QDialog.Accepted:
            return
        mappings[prop["path"]] = dlg.mapping()
    win.an_mappings = mappings
    _run(win, scan_files, win.an_paths, mappings, done=lambda res: _scanned(win, res))


def _scanned(win, res):
    raw, pair = res["raw"], None            # сравнение: первая часть → последняя
    _run(win, make_report, analysis_report_path(), raw, auto_mode(raw), pair, win.an_mappings,
         done=lambda r: _done(win, r, r["data"]))


def _done(win, res, data):
    lines, chips, errors = result_view(res, data)
    show_result(win, True, "✓  Анализ готов", lines, path=res["path"], chips=chips, after=errors)
    win.result_inn_box.setVisible(True)
    win.back_page = ANALYSIS_PAGE


def inn_from_result(win):
    """ИНН с экрана «Анализ готов»: тот же поиск, что на экране анализа (там виден ход работы)."""
    win.an_inn.setText(win.result_inn.text())
    win.pages.setCurrentIndex(ANALYSIS_PAGE)
    export_payments(win)


def export_payments(win):
    inn = clean_inn(win.an_inn.text())
    if win.worker:
        return
    if not inn.isdigit():
        QMessageBox.information(win, "Всё по ИНН", "Введите ИНН цифрами, например 040000011.")
        return
    _run(win, write_inn_payments, inn_report_path(inn), win.an_paths, inn, win.an_mappings,
         done=lambda res: _payments_done(win, res))


def _payments_done(win, res):
    if not res["path"]:
        QMessageBox.information(win, "Всё по ИНН", f"ИНН {res['inn']} не найден ни в одном из выбранных файлов.")
        return
    pays = [f"Платил: <b>{_int(res['paid'])}</b> платежей на <b>{format_money(res['paid_sum'])}</b>",
            f"Получал: <b>{_int(res['got'])}</b> платежей на <b>{format_money(res['got_sum'])}</b>"] if res["registries"] else []
    show_result(win, True, "✓  Всё по ИНН — готово", [
        f"Создан файл: <b>{os.path.basename(res['path'])}</b>",
        f"<span style='color:#64748B'>Папка: {os.path.dirname(res['path'])}</span>", "",
        f"ИНН <b>{res['inn']}</b>",
        f"В таблицах: <b>{_int(res['tables'])}</b> строк (лист «Из таблиц» — с подписью файла, листа и столбца)"] + pays,
        path=res["path"], note="Исходные файлы не изменены.")
    win.back_page = ANALYSIS_PAGE


def back(win):
    if not win.worker:
        win.pages.setCurrentIndex(0)
