"""Карточка файла на экране глубокого анализа: лист и колонки, вписанные от руки (C, C5, C5-C300, «G, H»).
У каждого файла — своя карточка; справа видно, готова ли она. Программа ничего не подставляет."""
import os

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QCheckBox, QComboBox, QFrame, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QVBoxLayout

from column_entry import MAIN, REGISTRY_FIELDS, to_mapping
from column_spec import SpecError
from ui_file_columns import formula_warning
from ui_file_widgets import collapsible, filter_edits

HINTS = {"inn": "напр. C или C5-C300", "name": "напр. B", "sums": "напр. G или G, H", "total": "напр. G400",
         "rec_inn": "напр. J", "rec": "напр. I", "date": "напр. B", "doc": "напр. A", "purpose": "напр. O"}
FLAGS = (("by_name", "ИНН нет — искать по названию (менее надёжно)"),
         ("repeats", "Одна организация может быть в нескольких строках (операции, выписка) — не ошибка"),
         ("separate", "Суммы не складывать — каждая колонка отдельно (например, План и Факт)"))
OK_STYLE, BAD_STYLE = "color:#15803D; font-weight:600;", "color:#C2410C; font-weight:600;"


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
        self.status.setWordWrap(True)
        self.status.setMaximumWidth(560)
        top.addWidget(self.status)
        lay.addLayout(top)
        if not self.sheets:                 # пустой файл — читать нечего
            self.use.setChecked(False)
            self.use.setEnabled(False)
            self.status.setText("пустой файл — пропускается")
            return
        self.use.setChecked(True)
        self.use.toggled.connect(self._check)
        self.edits, self.flags = {}, {}
        main = QGridLayout()
        self.sheet = QComboBox()
        for name, sh in self.sheets.items():
            self.sheet.addItem(name + ("   (скрытый лист)" if sh.get("hidden") else ""), name)
        self.sheet.currentIndexChanged.connect(self._check)
        self._fields(main, [("Лист", self.sheet)] + [(title, self._edit(key)) for key, title in MAIN])
        lay.addLayout(main)
        reg_title = QLabel("Реестр платежей — если в файле есть получатель:")
        reg_title.setObjectName("muted")
        lay.addWidget(reg_title)
        reg = QGridLayout()
        self._fields(reg, [(title, self._edit(key)) for key, title in REGISTRY_FIELDS])
        lay.addLayout(reg)
        more = collapsible("Дополнительно", lay)
        for key, text in FLAGS:
            self.flags[key] = QCheckBox(text)
            self.flags[key].toggled.connect(self._check)
            more.addRow("", self.flags[key])
        self.f_purpose, self.f_rec = filter_edits(more, self._check)
        self._check()

    @staticmethod
    def _fields(grid: QGridLayout, items: list):
        """Подписи сверху, поля под ними — в одну строку."""
        for col, (title, widget) in enumerate(items):
            label = QLabel(title)
            label.setObjectName("muted")
            grid.addWidget(label, 0, col)
            grid.addWidget(widget, 1, col)

    def _edit(self, key: str) -> QLineEdit:
        self.edits[key] = QLineEdit()
        self.edits[key].setPlaceholderText(HINTS[key])
        self.edits[key].textChanged.connect(self._check)
        return self.edits[key]

    def included(self) -> bool:
        return self.use.isChecked()

    def _build(self) -> dict:
        name = self.sheet.currentData()
        opts = {key: box.isChecked() for key, box in self.flags.items()}
        opts.update(f_purpose=self.f_purpose.text().strip(), f_rec=self.f_rec.text().strip())
        return to_mapping(name, self.sheets[name], {k: e.text().strip() for k, e in self.edits.items()}, opts)

    def problem(self) -> str:
        """Чего не хватает ("" — всё указано). Не взятый в анализ файл — без проблем."""
        if not self.included():
            return ""
        try:
            self._build()
            return ""
        except SpecError as e:
            return str(e)

    def _check(self):
        problem = self.problem()
        if not self.included():
            text, style = "не берётся в анализ", "color:#64748B;"
        elif problem:
            text, style = f"⚠ {problem}", BAD_STYLE
        else:
            m = self._build()
            note = formula_warning(m["sheet"], self.sheets[m["sheet"]], {"sums": m.get("sums") or m["cols"].get("amount") or []})
            text, style = (f"✓ готово · {note}" if note else "✓ готово"), (BAD_STYLE if note else OK_STYLE)
        self.status.setText(text)
        self.status.setStyleSheet(style)
        self.changed.emit()

    def mapping(self) -> list:
        """Для scan_files: настройка выбранного листа."""
        return [self._build()]
