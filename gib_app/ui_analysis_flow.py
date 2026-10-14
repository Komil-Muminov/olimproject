"""Глубокий анализ: чтение файлов в фоне, фильтры таблицы, «В Excel», платежи выбранного ИНН."""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QListWidgetItem, QMessageBox

from analysis_filter import Filters, apply as apply_filters
from analysis_model import ORG
from analysis_scan import AS_LIST, scan_files
from report_analysis import write_analysis_report
from report_inn import write_inn_payments
from report_paths import analysis_report_path, inn_report_path
from ui_column_dialog import ColumnDialog
from ui_result import show_result
from utils import to_number
from worker import Worker

ANALYSIS_PAGE = 3
ROLE_TITLES = {"плательщик": "Плательщики", "получатель": "Получатели", "организация": "Организации (списки)"}


def _int(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def _run(win, func, *args, done):
    for w in (win.an_filters, win.an_table, win.an_export_btn, win.an_pay_btn):
        w.setEnabled(False)
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
    ready = win.analysis is not None
    for w in (win.an_filters, win.an_table, win.an_export_btn):
        w.setEnabled(ready)
    win.an_pay_btn.setEnabled(ready and win.analysis.registry is not None)


def start(win, paths: list):
    win.analysis, win.an_paths, win.an_mappings, win.an_rows = None, list(paths), {}, []
    win.an_model.set_rows([], [], [], False)
    win.pages.setCurrentIndex(ANALYSIS_PAGE)
    _scan(win)


def _scan(win):
    win.an_info.setText("📚  " + ",  ".join(os.path.basename(p) for p in win.an_paths)
                        + "\nЧитаю… Большой реестр (800 тыс. строк) — примерно 15–20 секунд.")
    _run(win, scan_files, win.an_paths, win.an_mappings, done=lambda res: _scanned(win, res))


def _scanned(win, res):
    if "needs_mapping" in res:
        _ask_columns(win, res["needs_mapping"], res["guess"])
        return
    d = win.analysis = res["data"]
    win.an_parts.blockSignals(True)
    win.an_parts.clear()
    for p in d.parts:
        item = QListWidgetItem(f"{p.label}  ·  {p.source}")
        item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
        item.setCheckState(Qt.Checked)
        win.an_parts.addItem(item)
    win.an_parts.blockSignals(False)
    win.an_role.blockSignals(True)
    win.an_role.clear()
    win.an_role.addItem("Все", None)
    for r in d.roles:
        win.an_role.addItem(ROLE_TITLES[r], r)
    win.an_role.setVisible(len(d.roles) > 1)
    win.an_role.blockSignals(False)
    win.an_info.setText(f"📚  {',  '.join(os.path.basename(p) for p in d.paths)}  ·  частей: {len(d.parts)}  ·  "
                        f"организаций: {_int(len(d.entities))}.  Отметьте части и фильтры — таблица покажет нужное.")
    _idle(win)
    apply(win)


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
            return back(win)
        win.an_mappings[path] = dlg.mapping()
    elif q.clickedButton() is as_list:
        win.an_mappings[path] = AS_LIST
    else:
        return back(win)
    _scan(win)


def _filters(win) -> Filters:
    sel = [i for i in range(win.an_parts.count()) if win.an_parts.item(i).checkState() == Qt.Checked]
    return Filters(sel, win.an_role.currentData(), {n for n, b in win.an_status.items() if b.isChecked()},
                   {n for n, b in win.an_flags.items() if b.isChecked()},
                   to_number(win.an_min.text()) if win.an_min.text().strip() else None,
                   to_number(win.an_max.text()) if win.an_max.text().strip() else None, win.an_query.text())


def apply(win):
    d = win.analysis
    if d is None:
        return
    f = _filters(win)
    rows, st_counts, flag_counts = apply_filters(d, f)
    win.an_rows, win.an_sel, win.an_flags_on = rows, sorted(f.parts), f.flags
    win.an_model.set_rows(rows, [d.parts[i].label for i in win.an_sel], win.an_sel, len(d.roles) > 1)
    _widths(win, len(d.roles) > 1, len(win.an_sel))
    for chips, counts in ((win.an_status, st_counts), (win.an_flags, flag_counts)):
        for name, b in chips.items():
            b.setText(f"{name} ({_int(counts.get(name, 0))})")
    win.an_count.setText(f"Показано {_int(len(rows))} из {_int(len(d.entities))}"
                         + ("" if len(win.an_sel) > 1 else " · отметьте 2 части, чтобы видеть «что случилось»"))


def _widths(win, with_role: bool, parts: int):
    """Ширина колонок: ИНН, название, (кто), суммы по частям, всего, что случилось; «Ошибки» — остаток."""
    widths = [96, 300] + ([96] if with_role else []) + [140] * parts + [150, 112]
    for col, w in enumerate(widths):
        win.an_table.setColumnWidth(col, w)


def export(win):
    if win.analysis is None or win.worker:
        return
    _run(win, write_analysis_report, analysis_report_path(), win.analysis, win.an_rows, win.an_sel, win.an_flags_on,
         done=lambda res: _show(win, "✓  Сохранено в Excel", res["path"], f"Строк: <b>{_int(res['rows'])}</b>"))


def export_payments(win):
    rows = win.an_table.selectionModel().selectedRows() if win.analysis else []
    e = win.an_model.entity_at(rows[0].row()) if rows else None
    if not e or e.role == ORG or win.worker:
        QMessageBox.information(win, "Платежи", "Выберите в таблице строку плательщика или получателя из реестра.")
        return
    d = win.analysis
    _run(win, write_inn_payments, inn_report_path(e.inn), d.paths, e.inn, d.mappings,
         done=lambda res: _show(win, "✓  Платежи выгружены", res["path"],
                                f"Платил: <b>{res['paid']}</b> · получал: <b>{res['got']}</b> платежей"))


def _show(win, head, path, line):
    show_result(win, True, head, [f"Создан файл: <b>{os.path.basename(path)}</b>",
                                  f"<span style='color:#64748B'>Папка: {os.path.dirname(path)}</span>", "", line],
                path=path, note="Исходные файлы не изменены.")
    win.back_page = ANALYSIS_PAGE


def back(win):
    if not win.worker:
        win.analysis = None             # освободить память (таблица большая)
        win.pages.setCurrentIndex(0)
