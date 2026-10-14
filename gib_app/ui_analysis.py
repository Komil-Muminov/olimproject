"""Экран глубокого анализа: фильтры сверху, таблица организаций, «В Excel»."""
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView, QComboBox, QFrame, QGridLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QListWidget, QListView, QProgressBar, QPushButton, QSizePolicy, QTableView, QVBoxLayout, QWidget
)

import ui_analysis_flow as flow
from analysis_model import FLAGS, STATUSES
from ui_analysis_table import AnalysisTable


def _caption(text):
    lab = QLabel(text)
    lab.setObjectName("filterCaption")
    return lab


def _chips(win, names, attr):
    """Кнопки-переключатели (можно выбрать несколько); пусто — фильтр не действует."""
    box, lay = QWidget(), QHBoxLayout()
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(6)
    box.setLayout(lay)
    chips = {}
    for name in names:
        b = QPushButton(name)
        b.setObjectName("chip")
        b.setCheckable(True)
        b.setCursor(Qt.PointingHandCursor)
        b.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)      # не сжимать: текст со счётчиком виден целиком
        b.toggled.connect(lambda _: flow.apply(win))
        lay.addWidget(b)
        chips[name] = b
    lay.addStretch()
    setattr(win, attr, chips)
    return box


def build_analysis_card(win):
    card = QFrame()
    card.setObjectName("card")
    lay = QVBoxLayout(card)
    lay.setContentsMargins(22, 16, 22, 16)
    lay.setSpacing(8)
    win.an_info = QLabel()
    win.an_info.setObjectName("regInfo")
    win.an_info.setWordWrap(True)
    lay.addWidget(win.an_info)
    win.an_progress = QProgressBar()
    win.an_progress.setVisible(False)
    lay.addWidget(win.an_progress)

    win.an_filters = QWidget()
    grid = QGridLayout(win.an_filters)
    grid.setContentsMargins(0, 0, 0, 0)
    grid.setHorizontalSpacing(10)
    grid.setVerticalSpacing(6)
    win.an_parts = QListWidget()
    win.an_parts.setObjectName("partsList")
    win.an_parts.setFlow(QListView.LeftToRight)
    win.an_parts.setWrapping(True)
    win.an_parts.setMaximumHeight(64)
    win.an_parts.itemChanged.connect(lambda _: flow.apply(win))
    grid.addWidget(_caption("Части"), 0, 0)
    grid.addWidget(win.an_parts, 0, 1, 1, 3)

    win.an_role = QComboBox()
    win.an_role.currentIndexChanged.connect(lambda _: flow.apply(win))
    win.an_query = QLineEdit()
    win.an_query.setPlaceholderText("Поиск: ИНН или часть названия")
    win.an_query.setClearButtonEnabled(True)
    win.an_min, win.an_max = QLineEdit(), QLineEdit()
    win.an_min.setPlaceholderText("сумма от")
    win.an_max.setPlaceholderText("до")
    win.an_timer = QTimer(win)
    win.an_timer.setSingleShot(True)
    win.an_timer.setInterval(300)
    win.an_timer.timeout.connect(lambda: flow.apply(win))
    for edit in (win.an_query, win.an_min, win.an_max):
        edit.textEdited.connect(lambda _: win.an_timer.start())
    row = QHBoxLayout()
    row.addWidget(win.an_role)
    row.addWidget(win.an_query, 1)
    row.addWidget(win.an_min)
    row.addWidget(win.an_max)
    grid.addWidget(_caption("Кто / поиск"), 1, 0)
    grid.addLayout(row, 1, 1, 1, 3)
    grid.addWidget(_caption("Что случилось"), 2, 0)
    grid.addWidget(_chips(win, STATUSES, "an_status"), 2, 1, 1, 3)
    grid.addWidget(_caption("Ошибки"), 3, 0)
    grid.addWidget(_chips(win, FLAGS, "an_flags"), 3, 1, 1, 3)
    grid.setColumnStretch(1, 1)
    lay.addWidget(win.an_filters)

    win.an_model = AnalysisTable()
    t = win.an_table = QTableView()
    t.setObjectName("anTable")
    t.setModel(win.an_model)
    t.setSortingEnabled(True)
    t.verticalHeader().setVisible(False)
    t.verticalHeader().setDefaultSectionSize(26)
    t.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
    t.horizontalHeader().setStretchLastSection(True)
    t.setAlternatingRowColors(True)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setSelectionMode(QAbstractItemView.SingleSelection)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    lay.addWidget(t, 1)

    btns = QHBoxLayout()
    back = QPushButton("←  Назад")
    back.clicked.connect(lambda: flow.back(win))
    btns.addWidget(back)
    other = QPushButton("Другие файлы")
    other.clicked.connect(win.choose_file)
    btns.addWidget(other)
    win.an_count = QLabel()
    win.an_count.setObjectName("muted")
    btns.addWidget(win.an_count, 1)
    win.an_pay_btn = QPushButton("📄  Платежи выбранного ИНН")
    win.an_pay_btn.setToolTip("Все платежи выбранной в таблице организации — в Excel")
    win.an_pay_btn.clicked.connect(lambda: flow.export_payments(win))
    btns.addWidget(win.an_pay_btn)
    win.an_export_btn = QPushButton("📊  В Excel")
    win.an_export_btn.setObjectName("primary")
    win.an_export_btn.setToolTip("Сохранить в Excel то, что сейчас видно в таблице")
    win.an_export_btn.clicked.connect(lambda: flow.export(win))
    btns.addWidget(win.an_export_btn)
    lay.addLayout(btns)
    return card
