"""Строка листа в карточке файла: галочка «брать этот лист» и колонки ИНН и Сумма, вписанные от руки
(C, C5, C5-C300, «G, H»). Пустое поле — как у отмеченного листа выше: одинаковые колонки вписываются один раз."""
from PySide6.QtCore import Signal
from PySide6.QtWidgets import QCheckBox, QHBoxLayout, QLabel, QLineEdit, QWidget

FIELDS = (("inn", "ИНН"), ("sums", "Сумма"))
HINTS = {"inn": "напр. C или C5-C300", "sums": "напр. G или G, H"}
TIPS = {"inn": "Колонка с ИНН (обязательно). Строки берутся отсюда: C — вся колонка, C5-C300 — строки 5…300.",
        "sums": "Колонки с деньгами: G или G, H (в строке складываются) — для порога суммы и итогов."}
BAD_FIELD = "QLineEdit { border: 1px solid #DC2626; border-radius: 6px; }"
SHEET_WIDTH = 220


def header_row() -> QHBoxLayout:
    """Подписи над строками листов: «Лист | ИНН | Сумма» — под теми же ширинами."""
    lay = QHBoxLayout()
    for i, title in enumerate(["Лист"] + [t for _, t in FIELDS]):
        label = QLabel(title)
        label.setObjectName("muted")
        if i == 0:
            label.setFixedWidth(SHEET_WIDTH)
        lay.addWidget(label, 0 if i == 0 else 1)
    return lay


class SheetRow(QWidget):
    changed = Signal()

    def __init__(self, name: str, sheet: dict, checked: bool):
        super().__init__()
        self.name, self.sheet = name, sheet
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.use = QCheckBox(name + ("   (скрытый)" if sheet.get("hidden") else ""))
        self.use.setFixedWidth(SHEET_WIDTH)
        self.use.setToolTip(f"Отметьте, чтобы искать на листе «{name}»")
        self.use.setChecked(checked)
        self.use.toggled.connect(self._toggled)
        lay.addWidget(self.use)
        self.edits = {}
        for key, _ in FIELDS:
            self.edits[key] = QLineEdit()
            self.edits[key].textChanged.connect(lambda _: self.changed.emit())
            lay.addWidget(self.edits[key], 1)
        self._toggled()

    def included(self) -> bool:
        return self.use.isChecked()

    def typed(self) -> dict:
        return {k: e.text().strip() for k, e in self.edits.items()}

    def show_state(self, above: str, errors: dict):
        """above — с какого листа берутся пустые поля ("" — ни с какого); errors — поле → текст ошибки."""
        for key, edit in self.edits.items():
            edit.setPlaceholderText(f"как у «{above}»" if above and self.included() else HINTS[key])
            edit.setStyleSheet(BAD_FIELD if key in errors else "")
            edit.setToolTip(errors.get(key) or TIPS[key])

    def _toggled(self):
        for edit in self.edits.values():
            edit.setEnabled(self.included())
        self.changed.emit()
