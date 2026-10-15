"""Окно «Как читать файл» — перед анализом для каждого файла: лист, строка заголовков, что это (реестр
платежей или список организаций), где ИНН и какие столбцы суммировать. Догадка программы подставлена,
решает пользователь."""
import os

from openpyxl.utils import get_column_letter
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QHBoxLayout, QLabel, QListWidget,
    QListWidgetItem, QRadioButton, QSpinBox, QVBoxLayout, QWidget
)

from analysis_list_map import AS_LIST
from analysis_propose import REGISTRY, auto_columns
from constants import SUM_ROW_LABELS
from utils import format_money, norm_text

NONE = "— нет —"
FIELDS = {REGISTRY: [("payer_inn", "ИНН плательщика"), ("payer", "Плательщик"), ("rec_inn", "ИНН получателя"),
                     ("rec", "Получатель"), ("date", "Дата"), ("doc", "№ документа"), ("purpose", "Основание")],
          AS_LIST: [("inn", "ИНН  *"), ("name", "Название")]}


def text(r, i) -> str:
    """Ячейка для подписи: число — «1 234,56», текст — первые 30 знаков."""
    v = r[i] if i < len(r) else ""
    if isinstance(v, float):
        return format_money(v)
    return str(v).strip()[:30]


def _sample(rows, h: int) -> list:
    """Пример значений — первая строка под заголовками, кроме строк итогов («Ҳамагӣ», «Итого»)."""
    for r in rows[h + 1:]:
        if not any(isinstance(v, str) and norm_text(v).startswith(SUM_ROW_LABELS) for v in r):
            return r
    return []


class FileDialog(QDialog):
    def __init__(self, parent, prop: dict, k: int, n: int):
        super().__init__(parent)
        self.sheets, g = prop["sheets"], prop["guess"]
        self.setWindowTitle(f"Как читать файл ({k} из {n})")
        self.setMinimumWidth(640)
        lay = QVBoxLayout(self)
        name = QLabel(f"📄  <b>{os.path.basename(prop['path'])}</b>")
        lay.addWidget(name)
        form = QFormLayout()
        self.sheet = QComboBox()
        self.sheet.addItems(list(self.sheets))
        form.addRow("Лист", self.sheet)
        self.head = QSpinBox()
        form.addRow("Строка заголовков", self.head)
        kinds = QHBoxLayout()
        self.kind = QButtonGroup(self)
        for kind, title in ((REGISTRY, "Реестр платежей (плательщик → получатель)"), (AS_LIST, "Список организаций")):
            rb = QRadioButton(title)
            rb.setProperty("kind", kind)
            self.kind.addButton(rb)
            kinds.addWidget(rb)
        form.addRow("Что это", kinds)
        lay.addLayout(form)
        self.boxes, self.panels = {}, {}
        for kind, fields in FIELDS.items():
            panel, pf = QWidget(), QFormLayout()
            pf.setContentsMargins(0, 0, 0, 0)
            panel.setLayout(pf)
            for key, title in fields:
                self.boxes[kind, key] = QComboBox()
                self.boxes[kind, key].currentIndexChanged.connect(self._check)
                pf.addRow(title, self.boxes[kind, key])
            self.panels[kind] = panel
            lay.addWidget(panel)
        lay.addWidget(QLabel("Какие столбцы суммировать (можно несколько, например Дебет и Кредит):"))
        self.sums = QListWidget()
        self.sums.setMaximumHeight(170)
        self.sums.itemChanged.connect(self._check)
        lay.addWidget(self.sums)
        self.warn = QLabel()
        self.warn.setObjectName("warn")
        lay.addWidget(self.warn)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.button(QDialogButtonBox.Ok).setText("Готово")
        self.buttons.button(QDialogButtonBox.Cancel).setText("Отменить анализ")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        lay.addWidget(self.buttons)
        self._apply(g)
        self.sheet.currentTextChanged.connect(lambda _: self._apply(None))
        self.head.valueChanged.connect(lambda _: self._labels())
        self.kind.buttonToggled.connect(lambda _b, on: on and self._apply(None))

    # --- заполнение -------------------------------------------------------------------------------------
    def _rows(self) -> list:
        return self.sheets[self.sheet.currentText()]["rows"]

    def _apply(self, g):
        """Догадку (или авто-догадку по выбранному листу) — в поля окна."""
        if g:
            self.sheet.blockSignals(True)
            self.sheet.setCurrentText(g["sheet"])
            self.sheet.blockSignals(False)
        kind = g["kind"] if g else self.kind_now()
        g = g or dict(auto_columns(self._rows(), kind), kind=kind)
        sh = self.sheets[self.sheet.currentText()]
        self.head.blockSignals(True)
        self.head.setRange(sh["row0"] + 1, sh["row0"] + len(sh["rows"]))
        self.head.setValue(sh["row0"] + g["head"] + 1)
        self.head.blockSignals(False)
        for b in self.kind.buttons():
            b.blockSignals(True)
            b.setChecked(b.property("kind") == kind)
            b.blockSignals(False)
        self._labels()
        for (k, key), box in self.boxes.items():
            col = g["cols"].get(key) if k == kind else None
            box.setCurrentIndex(col + 1 if col is not None and col + 1 < box.count() else 0)
        for i in range(self.sums.count()):
            self.sums.item(i).setCheckState(Qt.Checked if i in g["sums"] else Qt.Unchecked)
        self._show_kind()

    def _labels(self):
        """Подписи столбцов «C: РМА (напр. 040058284)» по выбранной строке заголовков."""
        sh, rows = self.sheets[self.sheet.currentText()], self._rows()
        h = self.head.value() - sh["row0"] - 1
        heads, sample = rows[h], _sample(rows, h)
        width = max(len(r) for r in rows)
        names = [f"{get_column_letter(sh['col0'] + i + 1)}: {text(heads, i) or '(без названия)'}"
                 + (f"   (напр. {text(sample, i)})" if text(sample, i) else "") for i in range(width)]
        for box in self.boxes.values():
            idx = box.currentIndex()
            box.blockSignals(True)
            box.clear()
            box.addItems([NONE] + names)
            box.setCurrentIndex(max(idx, 0))
            box.blockSignals(False)
        checked = {i for i in range(self.sums.count()) if self.sums.item(i).checkState() == Qt.Checked}
        self.sums.blockSignals(True)
        self.sums.clear()
        for i, label in enumerate(names):
            item = QListWidgetItem(label)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if i in checked else Qt.Unchecked)
            self.sums.addItem(item)
        self.sums.blockSignals(False)
        self._check()

    def _show_kind(self):
        for kind, panel in self.panels.items():
            panel.setVisible(kind == self.kind_now())

    # --- результат --------------------------------------------------------------------------------------
    def kind_now(self) -> str:
        b = self.kind.checkedButton()
        return b.property("kind") if b else AS_LIST

    def _cols(self) -> dict:
        kind = self.kind_now()
        return {key: (box.currentIndex() - 1 if box.currentIndex() > 0 else None)
                for (k, key), box in self.boxes.items() if k == kind}

    def _sum_cols(self) -> list:
        return [i for i in range(self.sums.count()) if self.sums.item(i).checkState() == Qt.Checked]

    def _check(self):
        cols, kind = self._cols(), self.kind_now()
        if kind == REGISTRY:
            problem = "" if (cols.get("payer_inn") is not None or cols.get("rec_inn") is not None) and \
                self._sum_cols() else "Для реестра нужен хотя бы один ИНН и хотя бы один столбец суммы."
        else:
            problem = "" if cols.get("inn") is not None else "Укажите столбец с ИНН."
        self.buttons.button(QDialogButtonBox.Ok).setEnabled(not problem)
        self.warn.setText(problem)

    def mapping(self) -> dict:
        """Для scan_files: реестр — столбцы + "amount" (список сумм); список — "kind": "list" и "sums"."""
        sh, kind = self.sheets[self.sheet.currentText()], self.kind_now()
        m = {"sheet": self.sheet.currentText(), "head": self.head.value() - sh["row0"] - 1,
             "row0": sh["row0"], "col0": sh["col0"], "cols": self._cols()}
        if kind == REGISTRY:
            m["cols"]["amount"] = self._sum_cols()
            m["kind"] = REGISTRY
        else:
            m.update(kind=AS_LIST, sums=self._sum_cols())
        return m
