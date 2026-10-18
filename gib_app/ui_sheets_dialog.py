"""Окно «Листы и ИНН» — одно на все файлы: каждый лист каждого файла строкой, колонку ИНН программа находит
сама. Не нашла — пользователь вписывает колонку или отмечает «без ИНН» (лист попадёт в свод как есть)."""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QDialog, QDialogButtonBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget
)

from sheet_choice import choices, to_mappings

HEAD = ("Брать", "Файл · лист", "ИНН", "Название", "Без ИНН", "")
FOUND, TYPED, NOT_FOUND, AS_IS, SKIP = ("✓ нашла сама", "✓ вписано", "⚠ ИНН не найден — впишите колонку",
                                        "как есть", "не берём")
INFO = ("Программа нашла колонку ИНН на листах сама — проверьте. Где не нашла — впишите колонку (например, C) "
        "или отметьте «Без ИНН»: такой лист попадёт в свод как есть. Ненужный лист — снимите «Брать».")


def _centered(box: QCheckBox) -> QWidget:
    cell, lay = QWidget(), QHBoxLayout()
    lay.addWidget(box)
    lay.setAlignment(Qt.AlignCenter)
    lay.setContentsMargins(0, 0, 0, 0)
    cell.setLayout(lay)
    return cell


class SheetsDialog(QDialog):
    def __init__(self, parent, props: list):
        super().__init__(parent)
        self.props, self.rows = props, choices(props)
        self.auto = [r["inn"] for r in self.rows]          # что нашла программа — для пометки «нашла сама»
        self.setWindowTitle("Листы и ИНН")
        self.resize(860, 520)
        lay = QVBoxLayout(self)
        info = QLabel(INFO)
        info.setWordWrap(True)
        lay.addWidget(info)
        self.table = QTableWidget(len(self.rows), len(HEAD))
        self.table.setHorizontalHeaderLabels(HEAD)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionMode(QAbstractItemView.NoSelection)
        self.widgets = [self._row(i, r) for i, r in enumerate(self.rows)]
        head = self.table.horizontalHeader()
        head.setSectionResizeMode(1, QHeaderView.Stretch)
        head.setSectionResizeMode(5, QHeaderView.Stretch)
        lay.addWidget(self.table)
        self.warn = QLabel()
        self.warn.setObjectName("warn")
        self.warn.setWordWrap(True)
        lay.addWidget(self.warn)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.button(QDialogButtonBox.Ok).setText("Анализировать")
        self.buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        lay.addWidget(self.buttons)
        self._check()

    def _row(self, i: int, r: dict) -> dict:
        use, raw = QCheckBox(), QCheckBox()
        use.setChecked(True)
        inn, name = QLineEdit(r["inn"]), QLineEdit(r["name"])
        inn.setPlaceholderText("например C")
        name.setPlaceholderText("необязательно")
        title = QTableWidgetItem(f"{os.path.basename(r['path'])} · {r['sheet']}")
        title.setFlags(Qt.ItemIsEnabled)
        status = QTableWidgetItem()
        status.setFlags(Qt.ItemIsEnabled)
        self.table.setCellWidget(i, 0, _centered(use))
        self.table.setItem(i, 1, title)
        self.table.setCellWidget(i, 2, inn)
        self.table.setCellWidget(i, 3, name)
        self.table.setCellWidget(i, 4, _centered(raw))
        self.table.setItem(i, 5, status)
        for box in (use, raw):
            box.toggled.connect(self._check)
        for edit in (inn, name):
            edit.textChanged.connect(self._check)
        return {"use": use, "raw": raw, "inn": inn, "name": name, "status": status}

    def _check(self):
        for r, w in zip(self.rows, self.widgets):
            r.update(use=w["use"].isChecked(), raw=w["raw"].isChecked(), inn=w["inn"].text(), name=w["name"].text())
            for key in ("inn", "name", "raw"):
                w[key].setEnabled(r["use"] and (key == "raw" or not r["raw"]))
        self.chosen, problems = to_mappings(self.props, self.rows)
        bad = dict(problems)
        for i, (r, w) in enumerate(zip(self.rows, self.widgets)):
            text = self._status(i, r, bad)
            w["status"].setText(text)
            w["status"].setToolTip(text)                   # длинный текст обрезается — целиком при наведении
        ok = not problems and bool(self.chosen)
        self.buttons.button(QDialogButtonBox.Ok).setEnabled(ok)
        self.warn.setText(f"Не настроено листов: {len(problems)} — впишите колонку ИНН или отметьте «Без ИНН»."
                          if problems else "" if self.chosen else "Не выбрано ни одного листа.")

    def _status(self, i: int, r: dict, bad: dict) -> str:
        if not r["use"]:
            return SKIP
        if r["raw"]:
            return AS_IS
        if i in bad:
            return NOT_FOUND if not r["inn"].strip() else f"⚠ {bad[i]}"
        return FOUND if r["found"] and r["inn"].strip().upper() == self.auto[i] else TYPED

    def mappings(self) -> dict:
        """{путь: [настройки листов]} — только файлы, где выбран хоть один лист."""
        return self.chosen
