"""Главное окно: старт (ГИБ / Глубокий анализ), страница файла и результат."""
import os
from PySide6.QtCore import QSettings
from PySide6.QtWidgets import (
    QFileDialog, QHBoxLayout, QLabel, QMainWindow, QMessageBox,
    QPushButton, QStackedWidget, QVBoxLayout, QWidget
)

from constants import EXCEL_EXT
from styles import STYLE
import ui_flow
from ui_actions import (execute_check, execute_compare, execute_embed, load_plan,
                        render_found_info, update_option_marks)
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
        load_plan(self, path, ui_flow.current_mode(self))

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

    def _failed(self, err):
        self.worker = None
        self._busy(False)
        self.fill_progress.setVisible(False)
        self.update_fill_info()
        QMessageBox.critical(self, "Ошибка", err)

    def open_catalog(self):
        from catalog_dialog import CatalogDialog
        CatalogDialog(self).exec()
