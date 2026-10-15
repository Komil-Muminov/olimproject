"""Окно «Как читать файл» — перед анализом для каждого файла: лист и колонки, вписанные от руки
(C, C5, C5-C300, «G, H»). Программа ничего не подставляет — решает пользователь."""
import os

from PySide6.QtWidgets import QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QLabel, QLineEdit, QVBoxLayout

from column_entry import MAIN, REGISTRY_FIELDS, to_mapping
from column_spec import SpecError
from ui_file_columns import formula_warning
from ui_file_widgets import collapsible, dialog_buttons, filter_edits

HINTS = {"inn": "например C или C5-C300  (обязательно)", "name": "например B", "sums": "например G  или  G, H",
         "total": "ячейка с итогом в файле, например G400 — для сверки (необязательно)",
         "rec_inn": "например J", "rec": "например I", "date": "например B", "doc": "например A", "purpose": "например O"}
SYNTAX = ("C — вся колонка · C5 — с 5-й строки · C5-C300 — строки 5…300 · G, H — несколько колонок (суммы "
          "складываются). Строки берутся из поля ИНН.")


class FileDialog(QDialog):
    def __init__(self, parent, prop: dict, k: int, n: int):
        super().__init__(parent)
        self.sheets = prop["sheets"]
        self.setWindowTitle(f"Как читать файл ({k} из {n})")
        self.setMinimumWidth(620)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(f"📄  <b>{os.path.basename(prop['path'])}</b>"))
        form = QFormLayout()
        self.sheet = QComboBox()
        for name, sh in self.sheets.items():
            self.sheet.addItem(name + ("   (скрытый лист)" if sh.get("hidden") else ""), name)
        self.sheet.currentIndexChanged.connect(self._check)
        form.addRow("Лист", self.sheet)
        self.edits = {}
        for key, title in MAIN:
            form.addRow(title, self._edit(key))
        lay.addLayout(form)
        reg = collapsible("Реестр платежей — если в файле есть получатель", lay)
        for key, title in REGISTRY_FIELDS:
            reg.addRow(title, self._edit(key))
        more = collapsible("Дополнительно", lay)
        self.flags = {}
        for key, text in (("by_name", "ИНН нет — искать по названию (менее надёжно)"),
                          ("repeats", "Одна организация может быть в нескольких строках (операции, выписка) — не ошибка"),
                          ("separate", "Суммы не складывать — каждая колонка отдельно (например, План и Факт)")):
            self.flags[key] = QCheckBox(text)
            self.flags[key].toggled.connect(self._check)
            more.addRow("", self.flags[key])
        self.f_purpose, self.f_rec = filter_edits(more, self._check)
        hint = QLabel(SYNTAX)
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        lay.addWidget(hint)
        self.warn = QLabel()
        self.warn.setObjectName("warn")
        self.warn.setWordWrap(True)
        lay.addWidget(self.warn)
        self.buttons = dialog_buttons(self)
        lay.addWidget(self.buttons)
        self._check()

    def _edit(self, key: str) -> QLineEdit:
        self.edits[key] = QLineEdit()
        self.edits[key].setPlaceholderText(HINTS[key])
        self.edits[key].textChanged.connect(self._check)
        return self.edits[key]

    def _build(self) -> dict:
        name = self.sheet.currentData()
        opts = {key: box.isChecked() for key, box in self.flags.items()}
        opts.update(f_purpose=self.f_purpose.text().strip(), f_rec=self.f_rec.text().strip())
        return to_mapping(name, self.sheets[name], {k: e.text().strip() for k, e in self.edits.items()}, opts)

    def _check(self):
        try:
            m, problem = self._build(), ""
        except SpecError as e:
            m, problem = None, str(e)
        self.buttons.button(QDialogButtonBox.Ok).setEnabled(m is not None)
        note = formula_warning(m["sheet"], self.sheets[m["sheet"]], {"sums": m.get("sums") or m["cols"].get("amount") or []}) \
            if m else ""
        self.warn.setText(problem or note)

    def mapping(self) -> list:
        """Для scan_files: настройка выбранного листа."""
        return [self._build()]
