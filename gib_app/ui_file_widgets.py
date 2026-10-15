"""Мелкие виджеты окна «Как читать файл»: списки с галочками и кнопки «Готово / Отменить анализ»."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialogButtonBox, QLineEdit, QListWidget, QListWidgetItem


def checklist(height: int) -> QListWidget:
    w = QListWidget()
    w.setMaximumHeight(height)
    return w


def check_item(label: str, on: bool) -> QListWidgetItem:
    item = QListWidgetItem(label)
    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
    item.setCheckState(Qt.Checked if on else Qt.Unchecked)
    return item


def checked_rows(w: QListWidget) -> list:
    return [i for i in range(w.count()) if w.item(i).checkState() == Qt.Checked]


def dialog_buttons(dlg) -> QDialogButtonBox:
    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.button(QDialogButtonBox.Ok).setText("Готово")
    buttons.button(QDialogButtonBox.Cancel).setText("Отменить анализ")
    buttons.accepted.connect(dlg.accept)
    buttons.rejected.connect(dlg.reject)
    return buttons


def refill_combo(box, items: list):
    """Новые подписи пунктов, выбор (номер пункта) сохраняется."""
    idx = box.currentIndex()
    box.blockSignals(True)
    box.clear()
    box.addItems(items)
    box.setCurrentIndex(max(idx, 0))
    box.blockSignals(False)


def refill_checks(w: QListWidget, labels: list, checked: set):
    w.blockSignals(True)
    w.clear()
    for i, label in enumerate(labels):
        w.addItem(check_item(label, i in checked))
    w.blockSignals(False)


def filter_edits(panel_layout, on_change) -> tuple:
    """Отбор платежей реестра: основание / получатель содержат (слова через запятую). → (основание, получатель)."""
    out = []
    for title, hint in (("Основание содержит", "например: дивиденд, 30%  (пусто — все платежи)"),
                        ("Получатель содержит", "ИНН или часть названия, например: 010010058, бюджет")):
        edit = QLineEdit()
        edit.setPlaceholderText(hint)
        edit.textChanged.connect(on_change)
        panel_layout.addRow(title, edit)
        out.append(edit)
    return tuple(out)
