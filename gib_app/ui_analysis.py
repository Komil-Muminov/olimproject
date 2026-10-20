"""Экран глубокого анализа: выбранные файлы и кнопка «Анализ» внизу (как «Открыть в Excel» на экране итога)."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QProgressBar, QPushButton, QVBoxLayout

import ui_analysis_flow as flow

ANALYSIS_TIP = ("Excel, лист «Анализ»: каждая организация — строка, суммы по частям.\n"
                "Цвет строки — что случилось: Новый · Пропал · Вырос · Уменьшился · Без изменений · Появился и пропал.\n"
                "Оранжевая ячейка — ошибка: двойной платёж, неправильный ИНН, странная дата, повтор ИНН.\n"
                "В Excel есть легенда; наведите на цвет или заголовок — подсказка.")


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
    note = QLabel("Нажмите «Анализ»: для каждого файла откроется окно, где вы укажете колонки (ИНН, сумма…). "
                  "Исходные файлы не меняются · результат — в папке «Свод по глубокому анализу».")
    note.setObjectName("muted")
    note.setWordWrap(True)
    lay.addWidget(note)
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
    win.an_run_btn = QPushButton("📊  Анализ")
    win.an_run_btn.setObjectName("primary")
    win.an_run_btn.setToolTip(ANALYSIS_TIP)
    win.an_run_btn.setCursor(Qt.PointingHandCursor)
    win.an_run_btn.clicked.connect(lambda: flow.run(win))
    btns.addWidget(win.an_run_btn)
    lay.addLayout(btns)
    return card
