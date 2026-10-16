"""Окно «Что с чем сравнивать»: больше двух частей — пользователь выбирает «Было» и «Стало».
По умолчанию — первая и последняя часть (как без окна)."""
from PySide6.QtWidgets import QComboBox, QDialog, QDialogButtonBox, QFormLayout, QLabel, QVBoxLayout

from analysis_scan import preview_parts


def needs_pair(raw, mode: str) -> bool:
    """Спрашивать, только если есть выбор: частей больше двух и сравнение вообще будет
    (реестр + список не сравниваются — см. analysis_scan.build)."""
    reg, lists, _ = raw
    return len(preview_parts(reg, lists, mode)) > 2 and not (reg.files and lists.parts)


class PairDialog(QDialog):
    def __init__(self, parent, raw, mode: str):
        super().__init__(parent)
        reg, lists, _ = raw
        labels = [f"{p.label}  ·  {p.source}" for p in preview_parts(reg, lists, mode)]
        self.setWindowTitle("Что с чем сравнивать")
        self.setMinimumWidth(520)
        lay = QVBoxLayout(self)
        info = QLabel("Какие части сравнить? От этого зависят цвет строки («Что случилось») и «Изменение». "
                      "Остальные части тоже будут в таблице — со своими суммами.")
        info.setWordWrap(True)
        lay.addWidget(info)
        form = QFormLayout()
        self.was, self.now = QComboBox(), QComboBox()
        for box, index in ((self.was, 0), (self.now, len(labels) - 1)):
            box.addItems(labels)
            box.setCurrentIndex(index)
            box.currentIndexChanged.connect(lambda _: self._check())
        form.addRow("Было", self.was)
        form.addRow("Стало", self.now)
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

    def _check(self):
        """Одна и та же часть в «Было» и «Стало» — сравнивать нечего."""
        same = self.was.currentIndex() == self.now.currentIndex()
        self.buttons.button(QDialogButtonBox.Ok).setEnabled(not same)
        self.warn.setText("Выберите разные части." if same else "")

    def pair(self) -> tuple:
        """→ (было, стало) — номера частей в порядке времени."""
        return self.was.currentIndex(), self.now.currentIndex()
