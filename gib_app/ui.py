"""Главное окно приложения: выбор выписки, просмотр плана и запуск обработки."""
import os
from PySide6.QtCore import QSettings, Qt
from PySide6.QtWidgets import (
    QFileDialog, QHBoxLayout, QLabel, QMainWindow, QMessageBox,
    QPushButton, QStackedWidget, QVBoxLayout, QWidget
)

from constants import EXCEL_EXT
from planner import Plan
from styles import STYLE
from ui_actions import (
    execute_check, execute_compare, execute_embed, execute_run,
    populate_compare_options, render_found_info, update_option_marks
)
from ui_helpers import (
    format_done_summary, format_embed_summary, format_plan_description,
    format_warnings
)
from ui_views import build_found_card, build_result_card


class MainWindow(QMainWindow):
    def __init__(self, settings=None):
        super().__init__()
        self.setWindowTitle("Распределение по ГИБ")
        self.resize(900, 640)
        self.setAcceptDrops(True)
        self.settings = settings or QSettings("Olim", "GibDistributor")
        self.plan = self.worker = self.backup = self.result_path = None
        self._gib_text = ""

        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)
        lay = QVBoxLayout(root)
        lay.setContentsMargins(40, 30, 40, 30)
        lay.setSpacing(16)

        header_lay = QHBoxLayout()
        header_info = QVBoxLayout()
        title = QLabel("Распределение по ГИБ")
        title.setObjectName("appTitle")
        header_info.addWidget(title)
        sub = QLabel("Заполнение столбцов ГИБ по назначениям и сравнение списков по ИНН.")
        sub.setObjectName("muted")
        header_info.addWidget(sub)
        header_lay.addLayout(header_info)
        header_lay.addStretch()

        self.catalog_btn = QPushButton("📖  Справочник ГИБ")
        self.catalog_btn.setObjectName("secondary")
        self.catalog_btn.setToolTip("Просмотр и редактирование статей расходов бюджетной классификации")
        self.catalog_btn.clicked.connect(self.open_catalog)
        header_lay.addWidget(self.catalog_btn)
        lay.addLayout(header_lay)

        self.pages = QStackedWidget()
        lay.addWidget(self.pages, 1)

        self.drop = QPushButton("📂\n\nПеретащите сюда файл выписки\nили нажмите, чтобы выбрать")
        self.drop.setObjectName("dropZone")
        self.drop.clicked.connect(self.choose_file)
        self.pages.addWidget(self.drop)
        self.pages.addWidget(build_found_card(self))
        self.pages.addWidget(build_result_card(self))

    def open_file(self):
        if self.plan:
            self._start_file(self.plan.path)

    def _start_file(self, path):
        try:
            os.startfile(path)
        except OSError as e:
            QMessageBox.warning(self, "Не удалось открыть", f"{path}\n\n{e}")

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        files = [u.toLocalFile() for u in e.mimeData().urls() if u.toLocalFile().lower().endswith(EXCEL_EXT)]
        if files and not self.worker:
            self.load(files[0])

    def choose_file(self):
        if self.worker:
            return
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл выписки", self.settings.value("last_dir", ""),
                                              "Excel (*.xlsx *.xlsm *.xls)")
        if path:
            self.load(path)

    def load(self, path):
        self.settings.setValue("last_dir", os.path.dirname(path))
        try:
            plan = Plan(path)
        except Exception as ex:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл:\n{path}\n\n{ex}")
            return
        self.plan = plan
        self.file_label.setText("📄  " + os.path.basename(path))
        self.warn_text.setText("")
        lay = plan.layout
        gib = bool(lay or plan.header_only)
        self.embed_btn.setEnabled(gib)
        self.embed_btn.setVisible(gib)
        self.go_btn.setVisible(bool(lay))
        self.cmp_btn.setVisible(bool(plan.compare))
        self.chk_btn.setVisible(bool(plan.inn))
        self.note.setVisible(gib)

        text, enabled = format_plan_description(plan)
        if not gib and (plan.compare or plan.inn):
            self.found_text.setText("")
        else:
            self.found_text.setText(text)
            self.go_btn.setEnabled(enabled)
        self._gib_text = self.found_text.text()
        self.cmp_row.setVisible(bool(plan.compares))
        populate_compare_options(self)
        self.warn_text.setText(format_warnings(plan))
        self.pages.setCurrentIndex(1)

    def _compare_sheet_changed(self, i):
        if self.plan and 0 <= i < len(self.plan.compares):
            self.plan.compare = self.plan.compares[i]
            update_option_marks(self)
            render_found_info(self)

    def run_check(self):
        execute_check(self)

    def run_compare(self):
        execute_compare(self)

    def run(self):
        execute_run(self)

    def embed(self):
        execute_embed(self)

    def _busy(self, on):
        self.go_btn.setEnabled(not on and bool(self.plan and self.plan.items))
        self.embed_btn.setEnabled(not on)
        self.progress.setVisible(on)
        self.progress.setValue(0)

    def _embedded(self, res):
        self.worker = None
        self.progress.setVisible(False)
        self.result_path = res["path"]
        self._set_head(True, "✓  Автоматика встроена")
        lines, warn = format_embed_summary(res, self.plan)
        self.result_text.setText("<br>".join(lines))
        self.result_warn.setText(warn)
        self.backup = None
        self.backup_label.setText("")
        self.backup_btn.setVisible(False)
        self.pages.setCurrentIndex(2)

    def _set_head(self, ok, text):
        self.result_head.setText(text)
        self.result_head.setObjectName("resultOk" if ok else "resultBad")
        self.result_head.style().unpolish(self.result_head)
        self.result_head.style().polish(self.result_head)

    def _done(self, res):
        self.worker = None
        self.progress.setVisible(False)
        check = res["check"]
        self.result_path = self.plan.path
        self._set_head(check["ok"], "✓  Готово" if check["ok"] else "⚠  Готово, но итоги не сходятся")
        lines, warn = format_done_summary(res, self.plan)
        self.result_text.setText("<br>".join(lines))
        self.result_warn.setText("<br>".join(warn))
        self.backup = res.get("backup")
        self.backup_label.setText(f"🛡  Копия файла до изменений: {self.backup}" if self.backup else "")
        self.backup_btn.setVisible(bool(self.backup))
        self.pages.setCurrentIndex(2)

    def _failed(self, err):
        self.worker = None
        self._busy(False)
        QMessageBox.critical(self, "Ошибка", err)

    def open_catalog(self):
        from catalog_dialog import CatalogDialog
        CatalogDialog(self).exec()
