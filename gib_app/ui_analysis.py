"""Экран глубокого анализа: у каждого выбранного файла своя карточка с колонками (ui_file_card), внизу — «Анализ»."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QProgressBar, QPushButton, QScrollArea, QVBoxLayout, QWidget

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
