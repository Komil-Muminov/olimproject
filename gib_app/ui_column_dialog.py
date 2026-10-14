"""Окно «Укажите столбцы»: когда файл не узнан по заголовкам, пользователь сам выбирает, какой столбец
что значит. По умолчанию — для реестра платежей (сумма, ИНН…); для «Заполнить ГИБ» поля передаются свои."""
import os

from openpyxl.utils import get_column_letter
from PySide6.QtWidgets import QComboBox, QDialog, QDialogButtonBox, QFormLayout, QLabel, QVBoxLayout

from registry_columns import TITLES, is_enough

ORDER = ("amount", "payer_inn", "payer", "rec_inn", "rec", "doc", "date", "purpose")
NONE = "— нет —"
REGISTRY_NEED = "Обязательно: <b>Сумма</b> и хотя бы один <b>ИНН</b>."
REGISTRY_WARN = "Нужно указать «Сумму» и хотя бы один ИНН (плательщика или получателя)."


class ColumnDialog(QDialog):
    def __init__(self, parent, path: str, guess: dict, fields=ORDER, titles=TITLES, check=is_enough,
                 need=REGISTRY_NEED, warn=REGISTRY_WARN, required=("amount",)):
        super().__init__(parent)
        self.guess, self.check, self.warn_text = guess, check, warn
        self.setWindowTitle("Укажите столбцы")
        self.setMinimumWidth(560)
        lay = QVBoxLayout(self)
        info = QLabel(f"Файл «{os.path.basename(path)}», лист «{guess['sheet']}»: программа узнала не все столбцы.<br>"
                      f"Заголовки взяты из строки <b>{guess.get('row0', 0) + guess['head'] + 1}</b>. Укажите, какой "
                      f"столбец что значит. {need}")
        info.setWordWrap(True)
        lay.addWidget(info)
        col0 = guess.get("col0", 0)
        sample = guess.get("sample", [])
        example = lambda i: f"   (напр. {sample[i][:30]})" if i < len(sample) and sample[i] else ""
        names = [f"{get_column_letter(col0 + i + 1)}: {h or '(без названия)'}{example(i)}"
                 for i, h in enumerate(guess["heads"])]
        form = QFormLayout()
        self.boxes = {}
        for key in fields:
            box = QComboBox()
            box.addItems([NONE] + names)
            found = guess["cols"].get(key)
            box.setCurrentIndex(found + 1 if found is not None and found < len(names) else 0)
            box.currentIndexChanged.connect(self._check)
            form.addRow(titles[key] + ("  *" if key in required else ""), box)
            self.boxes[key] = box
        lay.addLayout(form)
        self.warn = QLabel()
        self.warn.setObjectName("warn")
        lay.addWidget(self.warn)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.button(QDialogButtonBox.Ok).setText("Готово")
        self.buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        lay.addWidget(self.buttons)
        self._check()

    def cols(self) -> dict:
        return {k: (b.currentIndex() - 1 if b.currentIndex() > 0 else None) for k, b in self.boxes.items()}

    def _check(self):
        ok = self.check(self.cols())
        self.buttons.button(QDialogButtonBox.Ok).setEnabled(ok)
        self.warn.setText("" if ok else self.warn_text)

    def mapping(self) -> dict:
        """Для read_registry: {"sheet", "head", "cols"}."""
        return {"sheet": self.guess["sheet"], "head": self.guess["head"], "cols": self.cols()}
