"""Мелкие виджеты окна «Как читать файл»: кнопки «Готово / Отменить анализ», поля отбора, сворачиваемый блок."""
from PySide6.QtWidgets import QDialogButtonBox, QFormLayout, QLineEdit, QToolButton, QWidget


def dialog_buttons(dlg) -> QDialogButtonBox:
    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.button(QDialogButtonBox.Ok).setText("Готово")
    buttons.button(QDialogButtonBox.Cancel).setText("Отменить анализ")
    buttons.accepted.connect(dlg.accept)
    buttons.rejected.connect(dlg.reject)
    return buttons


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


def collapsible(title: str, parent_layout) -> QFormLayout:
    """Сворачиваемый блок «▸ Заголовок»: → раскладка для его полей. По умолчанию свёрнут."""
    button, body = QToolButton(), QWidget()
    button.setText(f"▸  {title}")
    button.setCheckable(True)
    button.setStyleSheet("QToolButton { border: none; font-weight: 600; }")
    form = QFormLayout(body)
    form.setContentsMargins(16, 0, 0, 0)
    body.setVisible(False)
    button.toggled.connect(lambda on: (body.setVisible(on), button.setText(f"{'▾' if on else '▸'}  {title}")))
    parent_layout.addWidget(button)
    parent_layout.addWidget(body)
    return form
