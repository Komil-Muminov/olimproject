"""Экран глубокого анализа: у каждого выбранного файла своя карточка с колонками (ui_file_card), внизу — «Анализ»."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QLineEdit, QProgressBar, QPushButton, QScrollArea, QVBoxLayout, QWidget
)

import ui_analysis_flow as flow

ANALYSIS_TIP = ("Excel, лист «Анализ»: каждая организация — строка, суммы по частям.\n"
                "Цвет строки — что случилось: Новый · Пропал · Вырос · Уменьшился · Без изменений · Появился и пропал.\n"
                "Оранжевая ячейка — ошибка: двойной платёж, неправильный ИНН, странная дата, повтор ИНН.\n"
                "В Excel есть легенда; наведите на цвет или заголовок — подсказка.")
SYNTAX = ("Для каждого файла впишите колонки: C — вся колонка · C5 — с 5-й строки · C5-C300 — строки 5…300 · "
          "G, H — несколько колонок (суммы складываются). Строки берутся из поля ИНН. "
          "Исходные файлы не меняются · результат — в папке «Свод по глубокому анализу».")


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

    search = QHBoxLayout()                  # один ИНН — все его строки из всех файлов
    title = QLabel("🔎  Один ИНН во всех файлах:")
    title.setStyleSheet("font-weight:600;")
    search.addWidget(title)
    win.an_inn = QLineEdit()
    win.an_inn.setPlaceholderText("ИНН, например 040058284")
    win.an_inn.setMaxLength(20)
    win.an_inn.returnPressed.connect(lambda: flow.search_inn(win))
    search.addWidget(win.an_inn, 1)
    win.an_inn_btn = QPushButton("В Excel")
    win.an_inn_btn.setToolTip("Все строки этого ИНН из всех взятых файлов (колонки — из карточек), без предела по строкам")
    win.an_inn_btn.clicked.connect(lambda: flow.search_inn(win))
    search.addWidget(win.an_inn_btn)
    lay.addLayout(search)
    limit = QHBoxLayout()                   # порог суммы — и для «Анализа» (лист «Суммы …»), и для поиска по ИНН
    title = QLabel("💰  Порог суммы:   от")
    title.setStyleSheet("font-weight:600;")
    limit.addWidget(title)
    win.an_min, win.an_max = QLineEdit(), QLineEdit()
    for edit, hint in ((win.an_min, "напр. 2000"), (win.an_max, "без ограничения")):
        edit.setPlaceholderText(hint)
        edit.setMaximumWidth(220)
    limit.addWidget(win.an_min)
    limit.addWidget(QLabel("до"))
    limit.addWidget(win.an_max)
    note = QLabel("записи с такой суммой — отдельным листом; на итоги не влияет")
    note.setObjectName("muted")
    limit.addWidget(note)
    limit.addStretch()
    lay.addLayout(limit)

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
    win.an_run_btn = QPushButton("📊  Анализ")
    win.an_run_btn.setObjectName("primary")
    win.an_run_btn.setToolTip(ANALYSIS_TIP)
    win.an_run_btn.setCursor(Qt.PointingHandCursor)
    win.an_run_btn.clicked.connect(lambda: flow.run(win))
    btns.addWidget(win.an_run_btn)
    lay.addLayout(btns)
    return card
