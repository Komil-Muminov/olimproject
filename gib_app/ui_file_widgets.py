"""Мелкие виджеты окна «Как читать файл»: списки с галочками и кнопки «Готово / Отменить анализ»."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialogButtonBox, QListWidget, QListWidgetItem


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
