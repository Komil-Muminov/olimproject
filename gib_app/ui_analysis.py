"""Экран глубокого анализа: выбранные файлы и кнопка «Анализ»."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QProgressBar, QPushButton, QSizePolicy, QVBoxLayout
)

import ui_analysis_flow as flow

ANALYSIS_TIP = ("Excel, лист «Анализ»: каждая организация — строка, суммы по частям.\n"
                "Цвет строки — что случилось: Новый · Пропал · Вырос · Уменьшился · Без изменений · Появился и пропал.\n"
                "Оранжевая ячейка — ошибка: двойной платёж, неправильный ИНН, странная дата, повтор ИНН.\n"
                "В Excel есть легенда; наведите на цвет или заголовок — подсказка.")


def _action(win, title: str, sub: str, tip: str, slot):
    """Большая кнопка-карточка: название и одна строка пояснения."""
    b = QPushButton(f"{title}\n{sub}")
    b.setObjectName("actionCard")
    b.setToolTip(tip)
    b.setCursor(Qt.PointingHandCursor)
    b.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    b.clicked.connect(slot)
    return b


def build_analysis_card(win):
    card = QFrame()
    card.setObjectName("card")
    lay = QVBoxLayout(card)
    lay.setContentsMargins(28, 24, 28, 24)
    lay.setSpacing(14)
    win.an_files = QLabel()
    win.an_files.setObjectName("fileName")
    win.an_files.setWordWrap(True)
    lay.addWidget(win.an_files)

    row = QHBoxLayout()
    row.setSpacing(14)
    win.an_run_btn = _action(win, "📊  Анализ", "Excel: организации по частям, что случилось, ошибки",
                             ANALYSIS_TIP, lambda: flow.run(win))
    row.addWidget(win.an_run_btn)
    lay.addLayout(row)

    lay.addStretch()

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
    note = QLabel("Исходные файлы не меняются · результат — в папке «Свод по глубокому анализу»")
    note.setObjectName("muted")
    btns.addWidget(note)
    lay.addLayout(btns)
    return card
