"""Страница результата: одна функция показа для всех режимов (+ метки с подсказками при наведении)."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHBoxLayout, QLabel, QVBoxLayout, QWidget

CHIPS_PER_ROW = 4


def show_result(win, ok: bool, head: str, lines: list, warn: str = "", path=None, backup=None,
                note: str = "", chips=None):
    """ok — зелёный/красный заголовок; path — что откроет «Открыть в Excel»;
    backup — копия до изменений (кнопка «Открыть копию»), note — серая строка внизу;
    chips — [(заголовок, [(текст, цвет, подсказка)])] — метки под текстом."""
    win.worker = None
    win.result_path = path
    win.result_head.setText(head)
    win.result_head.setObjectName("resultOk" if ok else "resultBad")
    win.result_head.style().unpolish(win.result_head)
    win.result_head.style().polish(win.result_head)
    win.result_text.setText("<br>".join(lines))
    _fill_chips(win.result_chips, chips or [])
    win.result_warn.setText(warn)
    win.backup = backup
    win.backup_label.setText(f"🛡  Копия файла до изменений: {backup}" if backup else note)
    win.backup_btn.setVisible(bool(backup))
    win.pages.setCurrentIndex(2)


def _fill_chips(box: QWidget, groups: list):
    lay = box.layout()
    while lay.count():
        item = lay.takeAt(0)
        if item.widget():
            item.widget().deleteLater()
    for title, items in groups:
        cap = QLabel(title)
        cap.setObjectName("chipTitle")
        lay.addWidget(cap)
        for start in range(0, len(items), CHIPS_PER_ROW):
            row = QWidget()
            h = QHBoxLayout(row)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(8)
            for text, color, tip in items[start:start + CHIPS_PER_ROW]:
                chip = QLabel(text)
                chip.setObjectName("chip")
                chip.setStyleSheet(f"background: #{color};")
                chip.setToolTip(tip)
                chip.setCursor(Qt.WhatsThisCursor)
                h.addWidget(chip)
            h.addStretch()                  # метки — плотно слева
            lay.addWidget(row)
    box.setVisible(bool(groups))


def chips_box() -> QWidget:
    """Пустой контейнер для меток (кладётся на карточку результата)."""
    box = QWidget()
    lay = QVBoxLayout(box)
    lay.setContentsMargins(0, 4, 0, 0)
    lay.setSpacing(6)
    box.setVisible(False)
    return box
