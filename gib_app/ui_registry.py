"""Карточка «Реестры платежей»: итоги по файлам, «Свод по ИНН», поиск по ИНН."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QLineEdit, QProgressBar, QPushButton, QVBoxLayout
)

import ui_registry_flow as flow
from ui_registry_found import build_found_panel


def _label(name, rich=True):
    lab = QLabel()
    lab.setObjectName(name)
    lab.setWordWrap(True)
    if rich:
        lab.setTextFormat(Qt.RichText)
    return lab


def build_registry_card(win):
    card = QFrame()
    card.setObjectName("card")
    lay = QVBoxLayout(card)
    lay.setContentsMargins(28, 24, 28, 24)
    lay.setSpacing(12)

    win.reg_info = _label("regInfo")
    lay.addWidget(win.reg_info)
    win.reg_progress = QProgressBar()
    win.reg_progress.setVisible(False)
    lay.addWidget(win.reg_progress)

    row = QHBoxLayout()
    win.reg_inn = QLineEdit()
    win.reg_inn.setObjectName("bigSelect")
    win.reg_inn.setPlaceholderText("🔎  Поиск по ИНН: введите ИНН, например 040000011, и нажмите Enter")
    win.reg_inn.setMaxLength(20)
    win.reg_inn.setClearButtonEnabled(True)
    win.reg_inn.returnPressed.connect(lambda: flow.find(win))
    row.addWidget(win.reg_inn, 1)
    win.reg_find_btn = QPushButton("Найти")
    win.reg_find_btn.clicked.connect(lambda: flow.find(win))
    row.addWidget(win.reg_find_btn)
    lay.addLayout(row)

    lay.addWidget(build_found_panel(win), 1)
    win.reg_export_btn.clicked.connect(lambda: flow.export(win))
    lay.addStretch()

    note = QLabel("Исходные файлы не меняются: результат — новый Excel-файл в папке «Свод по глубокому анализу».")
    note.setObjectName("muted")
    lay.addWidget(note)
    btns = QHBoxLayout()
    back = QPushButton("←  Назад")
    back.clicked.connect(win.registry_back)
    btns.addWidget(back)
    other = QPushButton("Другие файлы")
    other.clicked.connect(win.choose_file)
    btns.addWidget(other)
    btns.addStretch()
    win.reg_report_btn = QPushButton("📊  Свод по ИНН")
    win.reg_report_btn.setObjectName("primary")
    win.reg_report_btn.setToolTip("Excel: плательщики и получатели по периодам, кто пропал и появился,\n"
                                  "резкий рост и падение, двойные платежи, неправильные ИНН, странные даты.")
    win.reg_report_btn.clicked.connect(lambda: flow.report(win))
    btns.addWidget(win.reg_report_btn)
    lay.addLayout(btns)
    return card
