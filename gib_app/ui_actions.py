"""«Встроить автоматику в Excel»: открыть файл, показать, что нашлось, создать копию с макросом."""
import os

from PySide6.QtWidgets import QFileDialog, QMessageBox

from excel_io import embed_automation
from planner import Plan
from ui_helpers import format_embed_summary, format_plan_description, format_warnings, template_path
from ui_result import show_result
from worker import Worker


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


def load_plan(win, path):
    """Файл для «Встроить автоматику»: что нашлось и можно ли встроить макрос."""
    win.settings.setValue("last_dir", os.path.dirname(path))
    try:
        plan = Plan(path)
    except Exception as ex:
        QMessageBox.critical(win, "Ошибка", f"Не удалось открыть файл:\n{path}\n\n{ex}")
        return
    win.plan = plan
    win.file_label.setText("📄  " + os.path.basename(path))
    gib = bool(plan.layout or plan.header_only)
    win.embed_btn.setEnabled(gib)
    win.embed_btn.setVisible(gib)
    win.note.setVisible(gib)
    win.found_text.setText(format_plan_description(plan)[0])
    win.warn_text.setText(format_warnings(plan))
    win.pages.setCurrentIndex(1)
