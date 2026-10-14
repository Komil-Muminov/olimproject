"""Глубокий анализ: файлы → «Анализировать» → чтение в фоне → Excel → итог простыми словами."""
import os

from PySide6.QtWidgets import QDialog, QMessageBox

from analysis_jobs import make_report
from analysis_list_map import FIELDS as LIST_FIELDS, TITLES as LIST_TITLES, list_ok
from analysis_scan import AS_LIST, auto_mode, scan_files
from report_inn import clean_inn, write_inn_payments
from report_paths import analysis_report_path, inn_report_path
from ui_column_dialog import ColumnDialog
from ui_analysis_result import result_view
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
    """«Анализ»: прочитать файлы в фоне, потом записать Excel."""
    if not win.worker:
        _run(win, scan_files, win.an_paths, win.an_mappings, done=lambda res: _scanned(win, res))


def _scanned(win, res):
    if "needs_mapping" in res:
        return _ask_columns(win, res["needs_mapping"], res["guess"])
    if "needs_inn" in res:
        return _ask_list(win, res["needs_inn"], res["guess"], "программа не нашла столбец с ИНН")
    if "needs_amount" in res:
        return _ask_list(win, res["needs_amount"], res["guess"], "несколько столбцов с числами — какой из них сумма?")
    raw, pair = res["raw"], None            # сравнение: первая часть → последняя
    mode = auto_mode(raw)
    _run(win, make_report, analysis_report_path(), raw, mode, pair, done=lambda r: _done(win, r, r["data"]))


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


def _ask_list(win, path: str, guess: dict, problem: str):
    """Список организаций: пользователь указывает столбец с ИНН (обязательно), название и сумму."""
    dlg = ColumnDialog(win, path, guess, LIST_FIELDS, LIST_TITLES, list_ok, need="Обязательно: <b>ИНН</b>. "
                       "Сумма — какой столбец считать деньгами («— нет —», если денег нет).",
                       warn="Укажите столбец с ИНН.", required=("inn",), problem=problem)
    if dlg.exec() != QDialog.Accepted:
        return
    win.an_mappings[path] = dict(dlg.mapping(), kind=AS_LIST)
    run(win)


def _done(win, res, data):
    lines, chips, errors = result_view(res, data)
    show_result(win, True, "✓  Анализ готов", lines, path=res["path"], chips=chips, after=errors)
    win.back_page = ANALYSIS_PAGE


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
