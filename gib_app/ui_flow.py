"""Стартовая страница: режимы и «Заполнить ГИБ» из двух файлов (выписка → ваш файл)."""
import os
from types import SimpleNamespace

from PySide6.QtWidgets import QFileDialog, QMessageBox

from excel_transfer import transfer_and_fill
from statement import read_statement, target_sheets
from ui_fill_columns import ask_statement_columns, ask_target_columns
from ui_helpers import format_done_summary, rows_list
from ui_result import show_result
from utils import format_money
from worker import Worker

EXCEL_FILTER = "Excel (*.xlsx *.xlsm *.xls)"


def current_mode(win) -> str:
    if win.mode_group.checkedId() == 1:
        return "analysis"
    return "fill" if win.fill_radio.isChecked() else "embed"


def update_start(win):
    mode = current_mode(win)
    win.gib_box.setVisible(mode != "analysis")
    win.fill_box.setVisible(mode == "fill")
    win.drop.setVisible(mode != "fill")
    win.drop.setText("📂\n\nПеретащите сюда файл\nили нажмите, чтобы выбрать\n\n" + (
        "Программа встроит макрос: дальше ГИБ проставятся сами" if mode == "embed"
        else "Любые файлы с ИНН, можно несколько — программа покажет таблицу с фильтрами"))


def pick(win, title):
    path, _ = QFileDialog.getOpenFileName(win, title, win.settings.value("last_dir", ""), EXCEL_FILTER)
    if path:
        win.settings.setValue("last_dir", os.path.dirname(path))
    return path


def set_statement(win, path):
    mapping = None
    try:
        sheet, pays = read_statement(path)
        if sheet is None:                       # столбцы не узнаны — пусть укажет пользователь
            mapping = ask_statement_columns(win, path)
            if mapping:
                _, pays = read_statement(path, mapping)
    except Exception as e:
        QMessageBox.critical(win, "Ошибка", f"Не удалось открыть выписку:\n{path}\n\n{e}")
        return
    win.st_path, win.st_pays, win.st_mapping = path, pays, mapping
    win.st_label.setText(os.path.basename(path))
    update_fill_info(win)


def set_target(win, path):
    cols = None
    try:
        sheets = target_sheets(path)
        if not sheets:                          # таблица не узнана и пустого листа нет — спросить
            picked = ask_target_columns(win, path)
            if picked:
                sheet, lay, cols = picked
                sheets = [(sheet, lay)]
    except Exception as e:
        QMessageBox.critical(win, "Ошибка", f"Не удалось открыть файл:\n{path}\n\n{e}")
        return
    win.tg_path, win.tg_sheets, win.tg_cols = path, sheets, cols
    win.tg_label.setText(os.path.basename(path))
    win.sheet_combo.blockSignals(True)
    win.sheet_combo.clear()
    win.sheet_combo.addItems([n for n, _ in sheets])
    win.sheet_combo.blockSignals(False)
    update_fill_info(win)


def update_fill_info(win):
    pays, sheets = win.st_pays, win.tg_sheets
    multi = len(sheets or []) > 1
    win.sheet_caption.setVisible(multi)
    win.sheet_combo.setVisible(multi)
    lines, ok = [], True
    if win.st_path:
        if pays:
            no_gib = sum(1 for p in pays if not p.gibs)
            lines.append(f"В выписке платежей: <b>{len(pays)}</b> на сумму <b>{format_money(sum(p.amount for p in pays))}</b>"
                         + (f"<span style='color:#B45309'> · без ГИБ: {no_gib}</span>" if no_gib else ""))
        else:
            lines.append("<span style='color:#DC2626'>В выписке не нашлось платежей "
                         "(нужны столбцы «Основание» и «Расходы»).</span>")
            ok = False
    if win.tg_path:
        if not sheets:
            lines.append("<span style='color:#DC2626'>Во втором файле нет таблицы со столбцом «Номгӯи корхонаҳо» / "
                         "«Основание» и нет пустого листа, где её можно создать.</span>")
            ok = False
        elif sheets[0][1] is None:
            lines.append(("Лист пустой" if multi else f"Лист «{sheets[0][0]}» пустой")
                         + " — программа создаст таблицу: р/т, Номгӯи корхонаҳо, № ҳуҷҷат, Сана, Ҳамагӣ, "
                           "Расходы, столбцы ГИБ и строку «Чамъ»")
        elif not multi:
            lines.append(f"Заполнится лист «{sheets[0][0]}»")
    if win.st_path and win.st_path == win.tg_path:
        lines.append("<span style='color:#DC2626'>Выберите разные файлы: сначала выписку, потом файл, куда заполнить.</span>")
        ok = False
    if ok and win.st_path and win.tg_path:
        lines.append("<span style='color:#64748B'>Допишутся после последней заполненной строки; "
                     "уже перенесённые раньше — пропускаются.</span>")
    win.fill_info.setText("<br>".join(lines))
    win.fill_btn.setEnabled(bool(ok and win.st_path and win.tg_path and not win.worker))


def run_fill(win):
    if not (win.st_path and win.tg_sheets) or win.worker:
        return
    name, lay = win.tg_sheets[max(0, win.sheet_combo.currentIndex())]
    win.fill_btn.setEnabled(False)
    win.fill_progress.setVisible(True)
    win.fill_progress.setValue(0)
    win.worker = Worker(transfer_and_fill, win.st_path, win.tg_path, name,
                        lay.header_row if lay else None, lay.purpose_col if lay else None, win.st_mapping, win.tg_cols)
    win.worker.progress.connect(lambda v, t: (win.fill_progress.setValue(v), win.fill_progress.setFormat(f"{t}  %p%")))
    win.worker.done.connect(win._filled)
    win.worker.failed.connect(win._failed)
    win.worker.start()


def show_filled(win, res):
    win.fill_progress.setVisible(False)
    lines = [f"Из выписки перенесено платежей: <b>{res['added']}</b> из {res['total']}"
             + (f" — строки {res['rows'][0]}–{res['rows'][1]}, лист «{res['sheet']}»" if res["rows"] else "")]
    if res["skipped"]:
        lines.append(f"Уже были в файле — пропущено: <b>{res['skipped']}</b>")
    if res["inserted"]:
        lines.append(f"В таблицу добавлено строк: <b>{res['inserted']}</b> (итоги «Чамъ» растянуты)")
    warn, ok = [], True
    fill = res["fill"]
    if fill:
        plan = SimpleNamespace(no_gib=res["no_gib"], mismatch=[], long_codes=res["long_codes"])
        more, warn = format_done_summary(fill, plan)
        lines += [""] + more
        ok = fill["check"]["ok"]
    elif res["no_gib"]:
        warn.append(f"⚠ Нет ГИБ в назначении — {rows_list(res['no_gib'])}. Заполните вручную.")
    head = "✓  Готово" if res["added"] else "✓  Новых платежей нет — всё уже перенесено"
    show_result(win, ok, head if ok else "⚠  Готово, но итоги не сходятся", lines, "<br>".join(warn),
                path=win.tg_path, backup=res.get("backup"))
    update_fill_info(win)
