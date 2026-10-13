"""Действия и обработчики операций интерфейса (запуск, проверки, сверка)."""
import os
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QMessageBox, QPushButton

from excel_io import embed_automation
from report import check_report_path, report_path, write_check_report, write_compare_report
from planner import Plan
from ui_helpers import format_embed_summary, format_plan_description, format_warnings, template_path
from ui_result import show_result
from worker import Worker


def populate_compare_options(win):
    for btn in win.sheet_group.buttons():
        win.sheet_group.removeButton(btn)
        btn.deleteLater()
    win.opt_texts = []
    for i, (name, d, a, b) in enumerate(win.plan.compares):
        what = (f"{a.where()[:1].upper() + a.where()[1:]}  ↔  {b.where()}\n"
                f"      кто появился, кто пропал и у кого изменилась сумма" if d.pairs else
                f"{a.where()[:1].upper() + a.where()[1:]}  ↔  {b.where()}\n"
                f"      какие организации есть только в одной части. В правой части нет сумм")
        win.opt_texts.append((name, what))
        btn = QPushButton()
        btn.setObjectName("sheetOption")
        btn.setCheckable(True)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setChecked(win.plan.compares[i] is win.plan.compare)
        win.sheet_group.addButton(btn, i)
        win.sheet_opts.addWidget(btn)
    update_option_marks(win)
    render_found_info(win)


def update_option_marks(win):
    for btn in win.sheet_group.buttons():
        name, what = win.opt_texts[win.sheet_group.id(btn)]
        btn.setText(f"{'●' if btn.isChecked() else '○'}   {name}\n      {what}")


def render_found_info(win):
    parts = [win._gib_text] if win._gib_text else []
    if win.plan.inn and win.mode == "analysis":
        parts.append(f"🧾 Таблиц с ИНН: <b>{len(win.plan.inn.tables)}</b>"
                     f"<span style='color:#64748B'> — кнопка «Проверить ИНН» найдёт дубликаты, "
                     f"разные названия одного ИНН и неправильные ИНН.</span>")
    win.found_text.setText("<br><br>".join(parts))
    win.found_text.setVisible(bool(parts))
    if not win.plan.compare:
        return
    sheet, d, a, b = win.plan.compare
    sheets = ("Итого · Новые · Пропавшие · Изменения · Без изменений · Дубликаты · Разные названия" if d.pairs and a.year and b.year
              else "Итого · Только справа · Только слева · " + ("Изменения · Без изменений" if d.pairs else "Есть в обоих")
              + " · Дубликаты · Разные названия")
    win.cmp_info.setText(f"Результат — новый Excel-файл, каждая категория на своём листе: {sheets}. Исходный файл не меняется.")


def execute_check(win):
    p = win.plan
    if not p or not p.inn:
        return
    try:
        path = write_check_report(check_report_path(p.path), p.path, p.inn)
    except Exception as e:
        QMessageBox.critical(win, "Ошибка", f"{type(e).__name__}: {e}")
        return
    r = p.inn
    show_result(win, True, "✓  Проверка готова — всё в порядке" if not r.problems else "✓  Проверка готова", [
        f"Создан файл: <b>{os.path.basename(path)}</b>", "",
        f"Проверено таблиц с ИНН: <b>{len(r.tables)}</b>",
        f"👥 Дубликаты ИНН: <b>{len(r.duplicates)}</b>",
        f"✏ Один ИНН — разные названия: <b>{len(r.names)}</b>",
        f"⚠ Неправильные ИНН: <b>{len(r.invalid)}</b>"], path=path, note="Исходный файл не изменён.")


def execute_compare(win):
    p = win.plan
    if not p or not p.compare:
        return
    sheet, d, a, b = p.compare
    try:
        path = write_compare_report(report_path(p.path, sheet), p.path, sheet, d, a, b)
    except Exception as e:
        QMessageBox.critical(win, "Ошибка", f"{type(e).__name__}: {e}")
        return
    show_result(win, True, "✓  Сравнение готово", [
        f"Создан файл: <b>{os.path.basename(path)}</b>", "",
        (f"🆕 Новые: <b>{len(d.only2)}</b>" if a.year and b.year else f"🆕 Только в правой части ({b.span}): <b>{len(d.only2)}</b>"),
        (f"❌ Пропавшие: <b>{len(d.only1)}</b>" if a.year and b.year else f"❌ Только в левой части ({a.span}): <b>{len(d.only1)}</b>"),
        *([f"🔄 Изменилась сумма: <b>{len(d.org_changed)}</b> орг.", f"✓ Без изменений: <b>{d.same}</b>"] if d.pairs
          else [f"✓ Есть в обоих списках: <b>{d.common}</b> (суммы не сравнивались)"]),
        f"👥 Дубликаты ИНН: <b>{len(d.duplicates)}</b>",
        f"✏ Разные названия при одном ИНН: <b>{len(d.org_names)}</b>"], path=path, note="Исходный файл не изменён.")


def execute_embed(win):
    p = win.plan
    lay = p.layout or p.header_only if p else None
    if not lay:
        return
    base = os.path.splitext(p.path)[0]
    dst, _ = QFileDialog.getSaveFileName(win, "Куда сохранить файл с автоматикой",
                                         f"{base} (автоматика).xlsm", "Книга Excel с макросами (*.xlsm)")
    if not dst:
        return
    if not dst.lower().endswith(".xlsm"):
        dst += ".xlsm"
    win._busy(True)
    win.worker = Worker(embed_automation, p.path, dst, template_path(), p.sheet, lay.header_row, lay.purpose_col)
    win.worker.progress.connect(lambda v, t: (win.progress.setValue(v), win.progress.setFormat(f"{t}  %p%")))
    win.worker.done.connect(lambda res: show_embedded(win, res))
    win.worker.failed.connect(win._failed)
    win.worker.start()


def show_embedded(win, res):
    win.progress.setVisible(False)
    lines, warn = format_embed_summary(res, win.plan)
    show_result(win, True, "✓  Автоматика встроена", lines, warn, path=res["path"])


def load_plan(win, path, mode):
    """Один файл: «Встроить автоматику» (mode="embed") или «Глубокий анализ» (mode="analysis")."""
    win.settings.setValue("last_dir", os.path.dirname(path))
    try:
        plan = Plan(path)
    except Exception as ex:
        QMessageBox.critical(win, "Ошибка", f"Не удалось открыть файл:\n{path}\n\n{ex}")
        return
    win.plan, win.mode = plan, mode
    win.file_label.setText("📄  " + os.path.basename(path))
    win.warn_text.setText("")
    analysis = mode == "analysis"
    gib = not analysis and bool(plan.layout or plan.header_only)
    win.embed_btn.setEnabled(gib)
    win.embed_btn.setVisible(gib)
    win.cmp_btn.setVisible(analysis and bool(plan.compare))
    win.chk_btn.setVisible(analysis and bool(plan.inn))
    win.note.setVisible(gib)
    win.cmp_row.setVisible(analysis and bool(plan.compares))
    if analysis:
        win._gib_text = "" if (plan.compare or plan.inn) else "В этом файле не нашлось таблиц с ИНН."
        populate_compare_options(win)
    else:
        win._gib_text = format_plan_description(plan)[0]
        win.found_text.setText(win._gib_text)
        win.found_text.setVisible(True)
        win.warn_text.setText(format_warnings(plan))
    win.pages.setCurrentIndex(1)
