"""Окно перед анализом: как разбить данные на части (по файлам / кварталам / месяцам) и что с чем
сравнивать («было → стало»)."""
from PySide6.QtWidgets import (
    QButtonGroup, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QHBoxLayout, QLabel, QRadioButton, QVBoxLayout
)

from analysis_periods import FILES, MODE_TITLES, available_modes
from analysis_scan import preview_parts


def needs_choice(raw) -> bool:
    """Спрашивать, только если есть выбор: больше двух частей или реестр можно разбить по периодам."""
    reg, lists, _ = raw
    return (len(available_modes(reg)) > 1 if reg.files else False) or len(preview_parts(reg, lists)) > 2


class CompareDialog(QDialog):
    def __init__(self, parent, raw):
        super().__init__(parent)
        self.reg, self.lists, _ = raw
        self.setWindowTitle("Что с чем сравнивать")
        self.setMinimumWidth(520)
        lay = QVBoxLayout(self)
        info = QLabel("Выберите, как разбить данные на части и какие части сравнить. От этого зависят «Что "
                      "случилось» (цвет строки) и «Изменение».")
        info.setWordWrap(True)
        lay.addWidget(info)
        form = QFormLayout()
        self.modes = QButtonGroup(self)
        row = QHBoxLayout()
        modes = available_modes(self.reg) if self.reg.files else [FILES]
        start = next((m for m in modes if len(preview_parts(self.reg, self.lists, m)) > 1), FILES)
        for mode in modes:
            rb = QRadioButton(MODE_TITLES[mode])
            rb.setProperty("mode", mode)
            rb.setChecked(mode == start)
            rb.toggled.connect(lambda on: on and self._fill())
            self.modes.addButton(rb)
            row.addWidget(rb)
        row.addStretch()
        if len(self.modes.buttons()) > 1:
            form.addRow("Части", row)
        self.was, self.now = QComboBox(), QComboBox()
        form.addRow("Было", self.was)
        form.addRow("Стало", self.now)
        self.form = form
        lay.addLayout(form)
        self.warn = QLabel()
        self.warn.setObjectName("warn")
        lay.addWidget(self.warn)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.button(QDialogButtonBox.Ok).setText("Дальше")
        self.buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        lay.addWidget(self.buttons)
        for box in (self.was, self.now):
            box.currentIndexChanged.connect(lambda _: self._check())
        self._fill()

    def mode(self) -> str:
        checked = self.modes.checkedButton()
        return checked.property("mode") if checked else FILES

    def _fill(self):
        labels = [f"{p.label}  ·  {p.source}" for p in preview_parts(self.reg, self.lists, self.mode())]
        for box, index in ((self.was, 0), (self.now, len(labels) - 1)):
            box.blockSignals(True)
            box.clear()
            box.addItems(labels)
            box.setCurrentIndex(index)
            box.blockSignals(False)
            self.form.setRowVisible(box, len(labels) > 1)
        self._check()

    def _single(self) -> bool:
        return self.was.count() < 2

    def _check(self):
        """Одна часть — сравнивать не с чем, просто «Дальше»."""
        same = not self._single() and self.was.currentIndex() == self.now.currentIndex()
        self.buttons.button(QDialogButtonBox.Ok).setEnabled(not same)
        self.warn.setText("Выберите разные части." if same else
                          "Одна часть — сравнения не будет." if self._single() else "")

    def choice(self) -> tuple:
        """→ (разбивка, (было, стало) или None, если часть одна)."""
        pair = None if self._single() else (self.was.currentIndex(), self.now.currentIndex())
        return self.mode(), pair
