"""Главное окно: старт (ГИБ / Глубокий анализ), страница файла и результат."""
import os
from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox, QStackedWidget, QVBoxLayout, QWidget

from constants import EXCEL_EXT
import ui_analysis_flow
import ui_flow
from ui_actions import execute_embed, load_plan
from ui_analysis import build_analysis_card
from ui_start import build_start_card
from ui_views import build_found_card, build_header, build_result_card


class MainWindow(QMainWindow):
    def __init__(self, settings=None):
        super().__init__()
        self.setWindowTitle("Распределение по ГИБ")
        avail = QApplication.primaryScreen().availableGeometry().height()
        self.resize(900, max(640, min(760, avail - 60)))       # повыше, если экран позволяет
        self.setAcceptDrops(True)
        self.settings = settings or QSettings("Olim", "GibDistributor")
        self.plan = self.worker = self.backup = self.result_path = None
        self.st_path = self.tg_path = self.st_pays = self.tg_sheets = self.st_mapping = self.tg_cols = None
        self.back_page = 0

        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)
        lay = QVBoxLayout(root)
        lay.setContentsMargins(40, 30, 40, 30)
        lay.setSpacing(16)

        lay.addLayout(build_header(self))

        self.pages = QStackedWidget()
        lay.addWidget(self.pages, 1)

        self.pages.addWidget(build_start_card(self))
        self.pages.addWidget(build_found_card(self))
        self.pages.addWidget(build_result_card(self))
        self.pages.addWidget(build_analysis_card(self))
        self.update_start()

    def update_start(self):
        ui_flow.update_start(self)

    def go_start(self):
        if not self.worker:
            self.pages.setCurrentIndex(self.back_page)
            self.back_page = 0

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
            self.open_files(files)

    def choose_file(self):
        if self.worker:
            return
        if ui_flow.current_mode(self) == "analysis":
            paths, _ = QFileDialog.getOpenFileNames(self, "Выберите файлы (можно несколько)",
                                                    self.settings.value("last_dir", ""), ui_flow.EXCEL_FILTER)
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Выберите файл", self.settings.value("last_dir", ""),
                                                  ui_flow.EXCEL_FILTER)
            paths = [path] if path else []
        if paths:
            self.open_files(paths)

    def open_files(self, paths):
        self.settings.setValue("last_dir", os.path.dirname(paths[0]))
        if ui_flow.current_mode(self) == "analysis":
            ui_analysis_flow.start(self, paths)
        else:
            self.load(paths[0])

    def load(self, path):
        load_plan(self, path)

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
