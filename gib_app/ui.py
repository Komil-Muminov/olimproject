"""Главное окно: старт (ГИБ / Глубокий анализ), страница файла и результат."""
import os
from PySide6.QtCore import QSettings
from PySide6.QtWidgets import (
    QFileDialog, QHBoxLayout, QLabel, QMainWindow, QMessageBox,
    QPushButton, QStackedWidget, QVBoxLayout, QWidget
)

from constants import EXCEL_EXT
from planner import Plan
from styles import STYLE
import ui_flow
from ui_actions import (execute_check, execute_compare, execute_embed, populate_compare_options,
                        render_found_info, update_option_marks)
from ui_helpers import format_embed_summary, format_plan_description, format_warnings
from ui_start import build_start_card
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
        self.mode = "fill"
        self.st_path = self.tg_path = self.st_pays = self.tg_sheets = None

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
        sub = QLabel("ГИБ — распределение сумм по назначениям. Глубокий анализ — сравнение и проверка по ИНН.")
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

        self.pages.addWidget(build_start_card(self))
        self.pages.addWidget(build_found_card(self))
        self.pages.addWidget(build_result_card(self))
        self.update_start()

    def update_start(self):
        ui_flow.update_start(self)

    def go_start(self):
        if not self.worker:
            self.pages.setCurrentIndex(0)

    def choose_statement(self):
        path = ui_flow.pick(self, "1. Выберите выписку")
        if path:
            ui_flow.set_statement(self, path)

    def choose_target(self):
        path = ui_flow.pick(self, "2. Выберите файл, куда заполнить")
        if path:
            ui_flow.set_target(self, path)

    def update_fill_info(self):
        ui_flow.update_fill_info(self)

    def run_fill(self):
        ui_flow.run_fill(self)

    def _filled(self, res):
        ui_flow.show_filled(self, res)

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
        if not files or self.worker:
            return
        if self.pages.currentIndex() == 0 and ui_flow.current_mode(self) == "fill":
            for f in files[:2]:     # первый — выписка, второй — куда заполнить
                (ui_flow.set_target if self.st_path and f != self.st_path else ui_flow.set_statement)(self, f)
        else:
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
        self.mode = ui_flow.current_mode(self)
        analysis = self.mode == "analysis"
        gib = not analysis and bool(plan.layout or plan.header_only)
        self.embed_btn.setEnabled(gib)
        self.embed_btn.setVisible(gib)
        self.cmp_btn.setVisible(analysis and bool(plan.compare))
        self.chk_btn.setVisible(analysis and bool(plan.inn))
        self.note.setVisible(gib)
        self.cmp_row.setVisible(analysis and bool(plan.compares))
        if analysis:
            self._gib_text = "" if (plan.compare or plan.inn) else "В этом файле не нашлось таблиц с ИНН."
            populate_compare_options(self)
        else:
            self._gib_text = format_plan_description(plan)[0]
            self.found_text.setText(self._gib_text)
            self.found_text.setVisible(True)
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

    def embed(self):
        execute_embed(self)

    def _busy(self, on):
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

    def _failed(self, err):
        self.worker = None
        self._busy(False)
        self.fill_progress.setVisible(False)
        self.update_fill_info()
        QMessageBox.critical(self, "Ошибка", err)

    def open_catalog(self):
        from catalog_dialog import CatalogDialog
        CatalogDialog(self).exec()
