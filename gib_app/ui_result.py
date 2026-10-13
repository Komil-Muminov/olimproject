"""Страница результата: одна функция показа для всех режимов."""


def show_result(win, ok: bool, head: str, lines: list, warn: str = "", path=None, backup=None,
                note: str = ""):
    """ok — зелёный/красный заголовок; path — что откроет «Открыть в Excel»;
    backup — копия до изменений (кнопка «Открыть копию»), note — серая строка внизу."""
    win.worker = None
    win.result_path = path
    win.result_head.setText(head)
    win.result_head.setObjectName("resultOk" if ok else "resultBad")
    win.result_head.style().unpolish(win.result_head)
    win.result_head.style().polish(win.result_head)
    win.result_text.setText("<br>".join(lines))
    win.result_warn.setText(warn)
    win.backup = backup
    win.backup_label.setText(f"🛡  Копия файла до изменений: {backup}" if backup else note)
    win.backup_btn.setVisible(bool(backup))
    win.pages.setCurrentIndex(2)
