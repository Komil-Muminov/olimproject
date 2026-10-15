"""Окно «Как читать файл» — перед анализом для каждого файла: какие листы читать (галочки) и для каждого
листа — строка заголовков, что это (реестр платежей или список), где ИНН и какие столбцы суммировать.
Догадка программы подставлена, решает пользователь."""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QHBoxLayout, QLabel, QRadioButton,
    QSpinBox, QVBoxLayout, QWidget
)

from analysis_list_map import AS_LIST
from analysis_propose import REGISTRY, auto_columns
from ui_file_columns import FIELDS, column_names, problem, sheet_view, to_mapping
from ui_file_widgets import check_item, checked_rows, checklist, dialog_buttons

NONE = "— нет —"


class FileDialog(QDialog):
    def __init__(self, parent, prop: dict, k: int, n: int):
        super().__init__(parent)
        self.sheets, g = prop["sheets"], prop["guess"]
        self.names = list(self.sheets)
        self.state = {g["sheet"]: {key: g.get(key) for key in ("head", "kind", "cols", "sums", "repeats")}}
        self.default_kind, self.cur = g["kind"], None
        self.setWindowTitle(f"Как читать файл ({k} из {n})")
        self.setMinimumWidth(680)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(f"📄  <b>{os.path.basename(prop['path'])}</b>"))
        lay.addWidget(QLabel("Листы — галочка: читать лист. Нажмите на лист, чтобы настроить его:"))
        self.sheet_list = checklist(110)
        for name in self.names:
            self.sheet_list.addItem(check_item(name, name == g["sheet"]))
        self.sheet_list.itemChanged.connect(self._check)
        lay.addWidget(self.sheet_list)
        form = self.form = QFormLayout()
        self.flip = QCheckBox("Организации идут по столбцам (таблица перевёрнута)")
        self.flip.toggled.connect(self._flipped)
        form.addRow("", self.flip)
        self.head = QSpinBox()
        self.head.valueChanged.connect(lambda _: self._labels())
        form.addRow("Строка заголовков", self.head)
        kinds, self.kind = QHBoxLayout(), QButtonGroup(self)
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
        self.by_name = QCheckBox("ИНН нет — искать по названию (менее надёжно)")
        self.by_name.toggled.connect(self._check)
        self.panels[AS_LIST].layout().addRow("", self.by_name)
        self.repeats = QCheckBox("Одна организация может быть в нескольких строках (операции, выписка) — не ошибка")
        self.panels[AS_LIST].layout().addRow("", self.repeats)
        lay.addWidget(QLabel("Какие столбцы суммировать (можно несколько, например Дебет и Кредит):"))
        self.sums = checklist(160)
        self.sums.itemChanged.connect(self._check)
        lay.addWidget(self.sums)
        self.separate = QCheckBox("Не складывать — каждый отмеченный столбец отдельно (например, План и Факт)")
        lay.addWidget(self.separate)
        self.warn = QLabel()
        self.warn.setObjectName("warn")
        lay.addWidget(self.warn)
        self.buttons = dialog_buttons(self)
        lay.addWidget(self.buttons)
        self.sheet_list.currentRowChanged.connect(self._switch)
        self.kind.buttonToggled.connect(lambda _b, on: on and self._load(self._guess(self.kind_now())))
        self.sheet_list.setCurrentRow(self.names.index(g["sheet"]))

    # --- лист: загрузить / сохранить ---------------------------------------------------------------------
    def _sh(self, name: str = None, flip: bool = None) -> dict:
        flip = self.flip.isChecked() if flip is None else flip
        return sheet_view(self.sheets[name or self.cur], flip)

    def _guess(self, kind: str, name: str = None, flip: bool = False) -> dict:
        return dict(auto_columns(self._sh(name, flip)["rows"], kind), kind=kind, transpose=flip)

    def _flipped(self, on: bool):
        """Перевёрнутая таблица — только список организаций."""
        self.kind.buttons()[0].setEnabled(not on)
        self.form.labelForField(self.head).setText("Столбец заголовков" if on else "Строка заголовков")
        self._load(self._guess(AS_LIST, flip=on))

    def _st(self, name: str) -> dict:
        """Настройка листа: что выбрал пользователь, а для неоткрытого листа — догадка программы."""
        return self.state.get(name) or self._guess(self.default_kind, name)

    def _switch(self, row: int):
        if self.cur:
            self.state[self.cur] = self._read()
        self.cur = self.names[row]
        self._load(self.state.get(self.cur) or self._guess(self.default_kind))

    def _load(self, st: dict):
        for w in [self.flip, self.head, self.by_name] + self.kind.buttons():
            w.blockSignals(True)
        self.flip.setChecked(bool(st.get("transpose")))
        self.by_name.setChecked(bool(st.get("by_name")))
        self.separate.setChecked(bool(st.get("separate")))
        self.repeats.setChecked(bool(st.get("repeats")))
        self.separate.setVisible(st["kind"] == AS_LIST)         # реестр — суммы всегда складываются
        self.kind.buttons()[0].setEnabled(not self.flip.isChecked())
        sh = self._sh()
        for w in [self.flip, self.by_name]:
            w.blockSignals(False)
        for w in [self.head] + self.kind.buttons():
            w.blockSignals(True)
        self.head.setRange(sh["row0"] + 1, sh["row0"] + len(sh["rows"]))
        self.head.setValue(sh["row0"] + st["head"] + 1)
        for b in self.kind.buttons():
            b.setChecked(b.property("kind") == st["kind"])
        for w in [self.head] + self.kind.buttons():
            w.blockSignals(False)
        self._labels()
        for (k, key), box in self.boxes.items():
            col = st["cols"].get(key) if k == st["kind"] else None
            box.setCurrentIndex(col + 1 if col is not None and col + 1 < box.count() else 0)
        for i in range(self.sums.count()):
            self.sums.item(i).setCheckState(Qt.Checked if i in st["sums"] else Qt.Unchecked)
        for kind, panel in self.panels.items():
            panel.setVisible(kind == st["kind"])
        self._check()

    def _labels(self):
        """Подписи столбцов по выбранной строке заголовков (выбор в полях и галочки сумм сохраняются)."""
        sh = self._sh()
        names = column_names(sh, self.head.value() - sh["row0"] - 1)
        checked = set(self._sum_cols())
        for box in self.boxes.values():
            idx = box.currentIndex()
            box.blockSignals(True)
            box.clear()
            box.addItems([NONE] + names)
            box.setCurrentIndex(max(idx, 0))
            box.blockSignals(False)
        self.sums.blockSignals(True)
        self.sums.clear()
        for i, label in enumerate(names):
            self.sums.addItem(check_item(label, i in checked))
        self.sums.blockSignals(False)
        self._check()

    def _read(self) -> dict:
        sh, kind = self._sh(), self.kind_now()
        return {"head": self.head.value() - sh["row0"] - 1, "kind": kind, "sums": self._sum_cols(),
                "transpose": self.flip.isChecked(), "by_name": kind == AS_LIST and self.by_name.isChecked(),
                "separate": kind == AS_LIST and self.separate.isChecked(),
                "repeats": kind == AS_LIST and self.repeats.isChecked(),
                "cols": {key: (b.currentIndex() - 1 if b.currentIndex() > 0 else None)
                         for (k, key), b in self.boxes.items() if k == kind}}

    # --- проверка и результат ----------------------------------------------------------------------------
    def kind_now(self) -> str:
        b = self.kind.checkedButton()
        return b.property("kind") if b else AS_LIST

    def _sum_cols(self) -> list:
        return checked_rows(self.sums)

    def _chosen(self) -> list:
        return [self.names[i] for i in checked_rows(self.sheet_list)]

    def _check(self):
        if not self.cur:
            return
        self.state[self.cur] = self._read()
        chosen = self._chosen()
        issues = [f"«{n}»: {problem(self._st(n))}" for n in chosen]
        issues = [x for x in issues if not x.endswith(": ")] or ([] if chosen else ["отметьте хотя бы один лист"])
        self.buttons.button(QDialogButtonBox.Ok).setEnabled(not issues)
        self.warn.setText("\n".join(issues))

    def mapping(self) -> list:
        """Для scan_files: по одной настройке на каждый отмеченный лист."""
        self.state[self.cur] = self._read()
        return [to_mapping(n, self._sh(n, bool(self._st(n).get("transpose"))), self._st(n)) for n in self._chosen()]
