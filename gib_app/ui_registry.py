"""Карточка «Реестры платежей»: итоги по файлам, «Свод по ИНН», поиск по ИНН."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QLineEdit, QProgressBar, QPushButton, QVBoxLayout
)

import ui_registry_flow as flow


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

    win.reg_files = _label("fileName", rich=False)
    lay.addWidget(win.reg_files)
    win.reg_info = _label("bigText")
    lay.addWidget(win.reg_info)
    win.reg_progress = QProgressBar()
    win.reg_progress.setVisible(False)
    lay.addWidget(win.reg_progress)

    title = QLabel("🔎  Поиск по ИНН")
    title.setObjectName("cmpTitle")
    lay.addWidget(title)
    row = QHBoxLayout()
    win.reg_inn = QLineEdit()
    win.reg_inn.setObjectName("bigSelect")
    win.reg_inn.setPlaceholderText("Введите ИНН, например 040000011")
    win.reg_inn.setMaxLength(20)
    win.reg_inn.returnPressed.connect(lambda: flow.find(win))
    row.addWidget(win.reg_inn, 1)
    win.reg_find_btn = QPushButton("Найти")
    win.reg_find_btn.clicked.connect(lambda: flow.find(win))
    row.addWidget(win.reg_find_btn)
    lay.addLayout(row)
    win.reg_found = _label("cmpInfo")
    lay.addWidget(win.reg_found)
    win.reg_export_btn = QPushButton("📄  Все платежи этого ИНН — в Excel")
    win.reg_export_btn.setObjectName("secondary")
    win.reg_export_btn.setVisible(False)
    win.reg_export_btn.clicked.connect(lambda: flow.export(win))
    lay.addWidget(win.reg_export_btn, 0, Qt.AlignLeft)
    lay.addStretch()

    note = QLabel("Исходные файлы не меняются: результат — новый Excel-файл в папке с реестрами.")
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
