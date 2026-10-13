"""Глубокий анализ реестров платежей: чтение в фоне, «Свод по ИНН», поиск и выгрузка по ИНН."""
import os

from PySide6.QtWidgets import QMessageBox

from registry_analysis import search
from registry_scan import scan_registries
from registry_stats import period_label
from report_inn import write_inn_payments
from report_paths import inn_report_path, registry_report_path
from report_registry import write_registry_report
from registry_index import clean_inn_query
from ui_registry_found import show_found
from ui_registry_inn import populate_inn_view, reset_inn_view
from ui_result import show_result
from utils import format_money
from worker import Worker

REGISTRY_PAGE = 3


def _int(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def _run(win, func, *args, done):
    win.reg_report_btn.setEnabled(False)
    win.reg_find_btn.setEnabled(False)
    win.reg_export_btn.setEnabled(False)
    win.reg_progress.setVisible(True)
    win.reg_progress.setValue(0)
    win.worker = Worker(func, *args)
    win.worker.progress.connect(lambda v, t: (win.reg_progress.setValue(v), win.reg_progress.setFormat(f"{t}  %p%")))
    win.worker.done.connect(done)
    win.worker.failed.connect(lambda err: (_idle(win), win._failed(err)))
    win.worker.start()


def _idle(win):
    win.worker = None
    win.reg_progress.setVisible(False)
    ready = win.registry is not None
    win.reg_report_btn.setEnabled(ready)
    win.reg_find_btn.setEnabled(ready)
    win.reg_export_btn.setEnabled(True)


def start_scan(win, paths: list):
    """Файлы из «Глубокого анализа»: реестры читаются в фоне; один обычный файл — по-старому."""
    win.registry, win.reg_paths = None, list(paths)
    reset_inn_view(win)
    win.reg_files.setText("📚  " + ",  ".join(os.path.basename(p) for p in paths))
    win.reg_info.setText("Читаю файлы… Большой реестр (800 тыс. строк) — примерно 15 секунд.")
    win.reg_export_btn.setVisible(False)
    win.pages.setCurrentIndex(REGISTRY_PAGE)
    _run(win, scan_registries, win.reg_paths, done=lambda res: _scanned(win, res))


def _scanned(win, res):
    win.worker = None
    if "not_registry" in res:
        path = res["not_registry"]
        _idle(win)
        if len(win.reg_paths) == 1:
            win.load(path)          # обычный файл: сравнение и проверка ИНН, как раньше
        else:
            QMessageBox.warning(win, "Не реестр", f"«{os.path.basename(path)}» — не реестр платежей: нет столбцов "
                                "«Плательщик», «ИНН Плат», «Получатель», «ИНН получ», «Сумма».\n\n"
                                "Несколько файлов можно выбрать только если все они — реестры.")
            win.pages.setCurrentIndex(0)
        return
    reg = win.registry = res["registry"]
    lines = [f"{reg.files[f].label}: <b>{_int(reg.files[f].rows)}</b> платежей на <b>{format_money(reg.files[f].total)}</b>"
             f"<span style='color:#64748B'> · {os.path.basename(reg.files[f].path)}</span>" for f in reg.order]
    lines.append(f"Всего: <b>{_int(sum(i.rows for i in reg.files))}</b> платежей на "
                 f"<b>{format_money(sum(i.total for i in reg.files))}</b> · плательщиков {_int(len(reg.payers))}, "
                 f"получателей {_int(len(reg.recs))}")
    win.reg_info.setText("<br>".join(lines))
    populate_inn_view(win, reg)
    _idle(win)


def report(win):
    reg = win.registry
    if not reg or win.worker:
        return
    months = sorted(m for i in reg.files for m in i.main_months)
    path = registry_report_path(period_label(months))
    _run(win, write_registry_report, path, reg, done=lambda res: _reported(win, res))


def _reported(win, res):
    _idle(win)
    c = res["counts"]
    lines = [f"Создан файл: <b>{os.path.basename(res['path'])}</b>",
             f"<span style='color:#64748B'>Папка: {os.path.dirname(res['path'])}</span>", ""]
    lines += [f"{k}: <b>{_int(v)}</b>" for k, v in c]
    show_result(win, True, "✓  Свод по ИНН готов", lines, path=res["path"], note="Исходные файлы не изменены.")
    win.back_page = REGISTRY_PAGE


def find(win, inn_val=None):
    reg = win.registry
    raw = (inn_val or win.reg_inn.text()).strip()
    if not reg or not raw:
        return
    inn = clean_inn_query(raw, reg)
    win.reg_inn.setText(inn)
    show_found(win, search(reg, inn), [reg.files[f].label for f in reg.order], raw)


def export(win):
    inn = getattr(win, "reg_found_inn", None)
    if not inn or win.worker:
        return
    path = inn_report_path(inn)
    _run(win, write_inn_payments, path, win.reg_paths, inn, done=lambda res: _exported(win, res))


def _exported(win, res):
    _idle(win)
    show_result(win, True, "✓  Платежи выгружены", [f"Создан файл: <b>{os.path.basename(res['path'])}</b>",
                f"<span style='color:#64748B'>Папка: {os.path.dirname(res['path'])}</span>", "",
                f"Платил: <b>{res['paid']}</b> платежей", f"Получал: <b>{res['got']}</b> платежей"],
                path=res["path"], note="Исходные файлы не изменены.")
    win.back_page = REGISTRY_PAGE
