"""Глубокий анализ: файлы → «Анализировать» → чтение в фоне → Excel → итог простыми словами."""
import os

from PySide6.QtWidgets import QDialog, QMessageBox

from analysis_model import FLAGS, STATUSES
from analysis_scan import AS_LIST, scan_files
from report_analysis import write_analysis_report
from report_inn import clean_inn, write_inn_payments
from report_pivot import write_pivot_report
from report_paths import analysis_report_path, inn_report_path, pivot_report_path
from ui_column_dialog import ColumnDialog
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
    for b in (win.an_run_btn, win.an_other_btn, win.an_pay_btn, win.an_pivot_btn):
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
    for b in (win.an_run_btn, win.an_other_btn, win.an_pay_btn, win.an_pivot_btn):
        b.setEnabled(True)


def run(win, kind: str = "report"):
    """kind: "report" — Excel «Анализ»; "pivot" — отдельный файл со сводной таблицей."""
    win.an_kind = kind
    if not win.worker:
        _run(win, scan_files, win.an_paths, win.an_mappings, done=lambda res: _scanned(win, res))


def _scanned(win, res):
    if "needs_mapping" in res:
        return _ask_columns(win, res["needs_mapping"], res["guess"])
    if win.an_kind == "pivot":
        return _run(win, write_pivot_report, pivot_report_path(), res["data"], done=lambda r: _pivot_done(win, r))
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
    run(win, win.an_kind)


def _pivot_done(win, res):
    show_result(win, True, "✓  Сводная таблица готова", [
        f"Создан файл: <b>{os.path.basename(res['path'])}</b>",
        f"<span style='color:#64748B'>Папка: {os.path.dirname(res['path'])}</span>", "",
        f"Лист «Сводная»: что случилось → организации ({_int(res['orgs'])}), по частям ({res['parts']}), "
        "сверху срезы — кнопки-фильтры.",
        f"Лист «Данные»: {_int(res['rows'])} строк — из них построена сводная."],
        path=res["path"], note="Исходные файлы не изменены.")
    win.back_page = ANALYSIS_PAGE


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


def export_payments(win):
    inn = clean_inn(win.an_inn.text())
    if win.worker:
        return
    if not inn.isdigit():
        QMessageBox.information(win, "Платежи", "Введите ИНН цифрами, например 040000011.")
        return
    _run(win, write_inn_payments, inn_report_path(inn), win.an_paths, inn, win.an_mappings,
         done=lambda res: _payments_done(win, res))


def _payments_done(win, res):
    if not res["path"]:
        QMessageBox.information(win, "Платежи", f"ИНН {res['inn']} не найден в реестрах ни среди плательщиков, "
                                "ни среди получателей.")
        return
    show_result(win, True, "✓  Платежи выгружены", [
        f"Создан файл: <b>{os.path.basename(res['path'])}</b>",
        f"<span style='color:#64748B'>Папка: {os.path.dirname(res['path'])}</span>", "",
        f"ИНН <b>{res['inn']}</b>",
        f"Платил: <b>{_int(res['paid'])}</b> платежей на <b>{format_money(res['paid_sum'])}</b>",
        f"Получал: <b>{_int(res['got'])}</b> платежей на <b>{format_money(res['got_sum'])}</b>"],
        path=res["path"], note="Исходные файлы не изменены.")
    win.back_page = ANALYSIS_PAGE


def back(win):
    if not win.worker:
        win.pages.setCurrentIndex(0)
