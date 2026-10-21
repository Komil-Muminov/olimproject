"""Карточка файла на экране фильтра: лист и колонки, вписанные от руки (C, C5, C5-C300, «G, H»): ИНН (какие строки
читать) и Сумма (для порога и итогов). Остальные колонки файла видны в отчёте как есть.
Справа — короткий статус; неверное поле обводится красным, полный текст ошибки — при наведении."""
import os

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QCheckBox, QComboBox, QFrame, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QVBoxLayout

from column_entry import to_mapping
from column_spec import SpecError, one_column, parse
from ui_file_columns import formula_warning

FIELDS = (("inn", "ИНН"), ("sums", "Сумма"))
HINTS = {"inn": "напр. C или C5-C300", "sums": "напр. G или G, H"}
TIPS = {"inn": "Колонка с ИНН (обязательно). Строки берутся отсюда: C — вся колонка, C5-C300 — строки 5…300.",
        "sums": "Колонки с деньгами: G или G, H (в строке складываются) — для порога суммы и итогов."}
OK_STYLE, BAD_STYLE, OFF_STYLE = "color:#15803D; font-weight:600;", "color:#C2410C; font-weight:600;", "color:#64748B;"
BAD_FIELD = "QLineEdit { border: 1px solid #DC2626; border-radius: 6px; }"
STATUS_WIDTH = 230


class FileCard(QFrame):
    changed = Signal()

    def __init__(self, prop: dict):
        super().__init__()
        self.setObjectName("actionBox")
        self.path, self.sheets = prop["path"], prop["sheets"]
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        top = QHBoxLayout()
        self.use = QCheckBox(f"📄  {os.path.basename(self.path)}")
        self.use.setStyleSheet("font-weight:700;")
        self.use.setToolTip("Снимите галочку, чтобы не брать этот файл в анализ")
        top.addWidget(self.use)
        top.addStretch()
        self.status = QLabel()
        self.status.setFixedWidth(STATUS_WIDTH)           # статус одной строкой — карточка не прыгает
        top.addWidget(self.status)
        lay.addLayout(top)
        if not self.sheets:                 # пустой файл — читать нечего
            self.use.setChecked(False)
            self.use.setEnabled(False)
            self._set_status("пустой файл — пропускается", OFF_STYLE)
            return
        self.use.setChecked(True)
        self.use.toggled.connect(self._check)
        self.edits = {}
        grid = QGridLayout()
        self.sheet = QComboBox()
        for name, sh in self.sheets.items():
            self.sheet.addItem(name + ("   (скрытый лист)" if sh.get("hidden") else ""), name)
        self.sheet.currentIndexChanged.connect(self._check)
        for col, (title, widget) in enumerate([("Лист", self.sheet)] + [(t, self._edit(k)) for k, t in FIELDS]):
            label = QLabel(title)
            label.setObjectName("muted")
            grid.addWidget(label, 0, col)
            grid.addWidget(widget, 1, col)
        lay.addLayout(grid)
        self._check()

    def _edit(self, key: str) -> QLineEdit:
        self.edits[key] = QLineEdit()
        self.edits[key].setPlaceholderText(HINTS[key])
        self.edits[key].setToolTip(TIPS[key])
        self.edits[key].textChanged.connect(self._check)
        return self.edits[key]

    def included(self) -> bool:
        return self.use.isChecked()

    def _build(self) -> dict:
        name = self.sheet.currentData()
        return to_mapping(name, self.sheets[name], {k: e.text().strip() for k, e in self.edits.items()})

    def problem(self) -> str:
        """Чего не хватает ("" — всё указано). Не взятый в анализ файл — без проблем."""
        if not self.included():
            return ""
        try:
            self._build()
            return ""
        except SpecError as e:
            return str(e)

    def _field_errors(self) -> dict:
        """Поле → текст ошибки (для красной рамки и подсказки)."""
        out = {} if self.edits["inn"].text().strip() else {"inn": "Обязательно: колонка с ИНН, например C или C5-C300"}
        for key, edit in self.edits.items():
            try:
                parse(edit.text()) if key == "sums" else one_column(edit.text(), dict(FIELDS)[key])
            except SpecError as e:
                out[key] = str(e)
        return out

    def _set_status(self, text: str, style: str, tip: str = ""):
        self.status.setText(text)
        self.status.setStyleSheet(style)
        self.status.setToolTip(tip or text)

    def _check(self):
        bad = self._field_errors() if self.included() else {}
        for key, edit in self.edits.items():
            edit.setStyleSheet(BAD_FIELD if key in bad else "")
            edit.setToolTip(bad.get(key) or TIPS[key])
        problem = self.problem()
        if not self.included():
            self._set_status("не берётся в анализ", OFF_STYLE)
        elif problem:
            self._set_status("⚠ проверьте поля", BAD_STYLE, problem)
        else:
            m = self._build()
            note = formula_warning(m["sheet"], self.sheets[m["sheet"]], {"sums": m.get("sums") or []})
            self._set_status("⚠ формулы без значений" if note else "✓ готово", BAD_STYLE if note else OK_STYLE, note)
        self.changed.emit()

    def mapping(self) -> list:
        """Настройка выбранного листа (список — на будущее несколько листов)."""
        return [self._build()]
