"""Карточка файла на экране фильтра: галочки по листам, у каждого отмеченного листа — колонки ИНН (какие строки
читать) и Сумма (для порога и итогов), вписанные от руки. Пустое поле — как у отмеченного листа выше.
Остальные колонки видны в отчёте как есть. Справа — короткий статус; полный текст ошибки — при наведении."""
import os

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QCheckBox, QFrame, QHBoxLayout, QLabel, QVBoxLayout

from column_entry import to_mapping
from column_spec import SpecError, one_column, parse
from ui_file_columns import formula_warning
from ui_sheet_row import FIELDS, SheetRow, header_row

OK_STYLE, BAD_STYLE, OFF_STYLE = "color:#15803D; font-weight:600;", "color:#C2410C; font-weight:600;", "color:#64748B;"
STATUS_WIDTH = 230


class FileCard(QFrame):
    changed = Signal()

    def __init__(self, prop: dict):
        super().__init__()
        self.setObjectName("actionBox")
        self.path, self.sheets, self.rows = prop["path"], prop["sheets"], []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        top = QHBoxLayout()
        self.use = QCheckBox(f"📄  {os.path.basename(self.path)}")
        self.use.setStyleSheet("font-weight:700;")
        self.use.setToolTip("Снимите галочку, чтобы не брать этот файл")
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
        lay.addLayout(header_row())
        for k, (name, sheet) in enumerate(self.sheets.items()):
            row = SheetRow(name, sheet, checked=k == 0)        # сначала отмечен первый лист
            row.use.setEnabled(len(self.sheets) > 1)            # единственный лист — всегда берётся
            row.changed.connect(self._check)
            lay.addWidget(row)
            self.rows.append(row)
        self._check()

    def included(self) -> bool:
        return self.use.isChecked()

    def _filled(self) -> list:
        """[(строка листа, что вписано с учётом листов выше, с какого листа взяты пустые поля)] — отмеченные листы."""
        out, last, source = [], {}, ""
        for row in self.rows:
            if not row.included():
                continue
            own = row.typed()
            typed = {k: own[k] or last.get(k, "") for k, _ in FIELDS}
            out.append((row, typed, source if any(not own[k] for k, _ in FIELDS) else ""))
            if any(own.values()):
                last, source = typed, row.name
        return out

    def problem(self) -> str:
        """Чего не хватает ("" — всё указано). Не взятый файл — без проблем."""
        if not self.included():
            return ""
        filled = self._filled()
        if not filled:
            return "Отметьте хотя бы один лист"
        for row, typed, _ in filled:
            try:
                to_mapping(row.name, row.sheet, typed)
            except SpecError as e:
                return f"Лист «{row.name}»: {e}" if len(self.rows) > 1 else str(e)
        return ""

    @staticmethod
    def _field_errors(typed: dict) -> dict:
        """Поле → текст ошибки (для красной рамки и подсказки)."""
        out = {} if typed["inn"] else {"inn": "Обязательно: колонка с ИНН, например C или C5-C300"}
        for key, title in FIELDS:
            try:
                parse(typed[key]) if key == "sums" else one_column(typed[key], title)
            except SpecError as e:
                out[key] = str(e)
        return out

    def _set_status(self, text: str, style: str, tip: str = ""):
        self.status.setText(text)
        self.status.setStyleSheet(style)
        self.status.setToolTip(tip or text)

    def _check(self):
        filled = {row: (typed, source) for row, typed, source in self._filled()} if self.included() else {}
        for row in self.rows:
            typed, source = filled.get(row, ({}, ""))
            row.show_state(source, self._field_errors(typed) if row in filled else {})
        problem = self.problem()
        if not self.included():
            self._set_status("не берётся", OFF_STYLE)
        elif problem:
            self._set_status("⚠ проверьте поля", BAD_STYLE, problem)
        else:
            notes = [n for m in self.mapping() if (n := formula_warning(m["sheet"], self.sheets[m["sheet"]], m))]
            self._set_status("⚠ формулы без значений" if notes else "✓ готово", BAD_STYLE if notes else OK_STYLE,
                             "\n".join(notes))
        self.changed.emit()

    def mapping(self) -> list:
        """Настройки отмеченных листов (по одной на лист)."""
        return [to_mapping(row.name, row.sheet, typed) for row, typed, _ in self._filled()]
