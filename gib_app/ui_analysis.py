"""Экран фильтра: у каждого выбранного файла своя карточка (лист, ИНН, сумма), ниже — фильтр по ИНН и/или сумме
и кнопка «Показать в Excel». Анализ делает сам пользователь — программа только показывает строки по фильтру."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QLineEdit, QProgressBar, QPushButton, QScrollArea, QVBoxLayout, QWidget
)

import ui_analysis_flow as flow

SYNTAX = ("Для каждого файла впишите колонки: C — вся колонка · C5 — с 5-й строки · C5-C300 — строки 5…300 · "
          "G, H — несколько колонок (суммы складываются). Исходные файлы не меняются · результат — в папке "
          "«Свод по глубокому анализу».")


def _field(lay, title: str, hint: str, width: int = None) -> QLineEdit:
    label = QLabel(title)
    label.setStyleSheet("font-weight:600;")
    lay.addWidget(label)
    edit = QLineEdit()
    edit.setPlaceholderText(hint)
    if width:
        edit.setMaximumWidth(width)
    lay.addWidget(edit)
    return edit


def build_analysis_card(win):
    card = QFrame()
    card.setObjectName("card")
    lay = QVBoxLayout(card)
    lay.setContentsMargins(28, 24, 28, 24)
    lay.setSpacing(12)
    hint = QLabel(SYNTAX)
    hint.setObjectName("muted")
    hint.setWordWrap(True)
    lay.addWidget(hint)
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    holder = QWidget()
    win.an_cards = QVBoxLayout(holder)
    win.an_cards.setContentsMargins(0, 0, 0, 0)
    win.an_cards.setSpacing(12)
    win.an_cards.addStretch()
    scroll.setWidget(holder)
    lay.addWidget(scroll, 1)

    filters = QHBoxLayout()                 # фильтр: ИНН (в любой колонке) и/или сумма строки от … до …
    win.an_inn = _field(filters, "🔎  ИНН:", "напр. 040058284 (пусто — все)", 260)
    win.an_min = _field(filters, "   💰  Сумма от", "напр. 2000", 180)
    win.an_max = _field(filters, "до", "без ограничения", 180)
    filters.addStretch()
    lay.addLayout(filters)
    for edit in (win.an_inn, win.an_min, win.an_max):
        edit.returnPressed.connect(lambda: flow.run(win))

    win.an_progress = QProgressBar()
    win.an_progress.setVisible(False)
    lay.addWidget(win.an_progress)
    btns = QHBoxLayout()
    back = QPushButton("←  Назад")
    back.clicked.connect(lambda: flow.back(win))
    btns.addWidget(back)
    win.an_other_btn = QPushButton("Другие файлы")
    win.an_other_btn.clicked.connect(win.choose_file)
    btns.addWidget(win.an_other_btn)
    btns.addStretch()
    win.an_run_btn = QPushButton("🔎  Показать в Excel")
    win.an_run_btn.setObjectName("primary")
    win.an_run_btn.setToolTip("Строки из всех взятых файлов по фильтру — как есть, со всеми колонками")
    win.an_run_btn.setCursor(Qt.PointingHandCursor)
    win.an_run_btn.clicked.connect(lambda: flow.run(win))
    btns.addWidget(win.an_run_btn)
    lay.addLayout(btns)
    return card
