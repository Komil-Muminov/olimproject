"""Диалоговое окно управления справочником кодов ГИБ (просмотр, добавление, редактирование, удаление)."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QFormLayout, QHBoxLayout, QHeaderView, QInputDialog,
    QLabel, QLineEdit, QMessageBox, QPushButton, QTableWidget,
    QTableWidgetItem, QVBoxLayout
)

from catalog_manager import GibCatalogManager


class CatalogItemDialog(QDialog):
    """Окно ввода/редактирования кода и наименования статьи."""

    def __init__(self, parent=None, code="", name="", is_edit=False):
        super().__init__(parent)
        self.setWindowTitle("Редактировать статью" if is_edit else "Добавить статью ГИБ")
        self.resize(480, 160)
        lay = QVBoxLayout(self)

        form = QFormLayout()
        self.code_edit = QLineEdit(code)
        self.code_edit.setPlaceholderText("Например: 2211")
        if is_edit:
            self.code_edit.setReadOnly(True)
        self.name_edit = QLineEdit(name)
        self.name_edit.setPlaceholderText("Например: Хизматрасонии коммуналӣ")
        form.addRow("Код ГИБ:", self.code_edit)
        form.addRow("Наименование:", self.name_edit)
        lay.addLayout(form)

        btns = QHBoxLayout()
        btns.addStretch()
        ok_btn = QPushButton("Сохранить")
        ok_btn.setObjectName("primary")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        btns.addWidget(cancel_btn)
        btns.addWidget(ok_btn)
        lay.addLayout(btns)

    def get_data(self) -> tuple[str, str]:
        return self.code_edit.text().strip(), self.name_edit.text().strip()


class CatalogDialog(QDialog):
    """Главный диалог справочника кодов ГИБ."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Справочник статей ГИБ (Маълумотнома)")
        self.resize(720, 520)
        self.manager = GibCatalogManager()

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)

        top = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("🔍  Поиск по коду или названию статьи...")
        self.search.textChanged.connect(self.refresh_table)
        top.addWidget(self.search)
        root.addLayout(top)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Код ГИБ", "Наименование статьи расходов"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.doubleClicked.connect(self.edit_selected)
        root.addWidget(self.table)

        actions = QHBoxLayout()
        add_btn = QPushButton("➕ Добавить")
        add_btn.setObjectName("primary")
        add_btn.clicked.connect(self.add_item)
        actions.addWidget(add_btn)

        edit_btn = QPushButton("✏️ Изменить")
        edit_btn.clicked.connect(self.edit_selected)
        actions.addWidget(edit_btn)

        del_btn = QPushButton("🗑️ Удалить")
        del_btn.clicked.connect(self.delete_selected)
        actions.addWidget(del_btn)

        reset_btn = QPushButton("🔄 Сброс по умолчанию")
        reset_btn.setObjectName("secondary")
        reset_btn.clicked.connect(self.reset_defaults)
        actions.addWidget(reset_btn)

        actions.addStretch()
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        actions.addWidget(close_btn)
        root.addLayout(actions)

        self.refresh_table()

    def refresh_table(self):
        query = self.search.text().strip().lower()
        items = self.manager.get_all()
        filtered = [(k, v) for k, v in items.items() if not query or query in k.lower() or query in v.lower()]
        self.table.setRowCount(len(filtered))
        for r, (code, name) in enumerate(filtered):
            item_code = QTableWidgetItem(code)
            item_code.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(r, 0, item_code)
            self.table.setItem(r, 1, QTableWidgetItem(name))

    def add_item(self):
        dlg = CatalogItemDialog(self)
        if dlg.exec() == QDialog.Accepted:
            code, name = dlg.get_data()
            if not code or not name:
                QMessageBox.warning(self, "Ошибка", "Код и наименование обязательны для заполнения.")
                return
            self.manager.set(code, name)
            self.refresh_table()

    def edit_selected(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Внимание", "Выберите статью для редактирования.")
            return
        code = self.table.item(row, 0).text()
        name = self.table.item(row, 1).text()
        dlg = CatalogItemDialog(self, code=code, name=name, is_edit=True)
        if dlg.exec() == QDialog.Accepted:
            _, new_name = dlg.get_data()
            if not new_name:
                QMessageBox.warning(self, "Ошибка", "Наименование не может быть пустым.")
                return
            self.manager.set(code, new_name)
            self.refresh_table()

    def delete_selected(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Внимание", "Выберите статью для удаления.")
            return
        code = self.table.item(row, 0).text()
        name = self.table.item(row, 1).text()
        reply = QMessageBox.question(self, "Подтверждение", f"Удалить статью ГИБ {code} «{name}»?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.manager.delete(code)
            self.refresh_table()

    def reset_defaults(self):
        reply = QMessageBox.question(self, "Сброс справочника",
                                     "Восстановить стандартный справочник кодов ГИБ?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.manager.reset_to_default()
            self.refresh_table()
