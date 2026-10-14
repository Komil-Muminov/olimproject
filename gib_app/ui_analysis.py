"""Экран глубокого анализа: выбранные файлы и одна кнопка «Анализировать» — результат в Excel."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QProgressBar, QPushButton, QVBoxLayout

import ui_analysis_flow as flow

HOW = ("<b>Что будет в Excel</b><br>"
       "• Каждая организация (ИНН) — одна строка: суммы по каждой части (файл, период, таблица на листе) и всего.<br>"
       "• <b>Цвет строки</b> — что случилось между первой и последней частью: "
       "<span style='background:#DCFCE7'>&nbsp;Новый&nbsp;</span> "
       "<span style='background:#FEE2E2'>&nbsp;Пропал&nbsp;</span> "
       "<span style='background:#DBEAFE'>&nbsp;Вырос&nbsp;</span> "
       "<span style='background:#FEF9C3'>&nbsp;Уменьшился&nbsp;</span> "
       "Без изменений "
       "<span style='background:#EDE9FE'>&nbsp;Появился и пропал&nbsp;</span><br>"
       "• <b>Ошибки</b> — <span style='background:#FED7AA'>&nbsp;оранжевая ячейка&nbsp;</span>: двойной платёж, "
       "неправильный ИНН, странная дата, повтор ИНН. Подробно — на втором листе.<br>"
       "• Наведите на заголовок или цвет в легенде — появится подсказка. Отбирать — обычным фильтром Excel.")


def build_analysis_card(win):
    card = QFrame()
    card.setObjectName("card")
    lay = QVBoxLayout(card)
    lay.setContentsMargins(28, 24, 28, 24)
    lay.setSpacing(12)
    win.an_files = QLabel()
    win.an_files.setObjectName("fileName")
    win.an_files.setWordWrap(True)
    lay.addWidget(win.an_files)
    how = QLabel(HOW)
    how.setObjectName("regInfo")
    how.setWordWrap(True)
    how.setTextFormat(Qt.RichText)
    lay.addWidget(how)
    lay.addStretch()
    win.an_progress = QProgressBar()
    win.an_progress.setVisible(False)
    lay.addWidget(win.an_progress)
    note = QLabel("Исходные файлы не меняются. Excel сохранится в папку «Свод по глубокому анализу».")
    note.setObjectName("muted")
    lay.addWidget(note)
    btns = QHBoxLayout()
    back = QPushButton("←  Назад")
    back.clicked.connect(lambda: flow.back(win))
    btns.addWidget(back)
    win.an_other_btn = QPushButton("Другие файлы")
    win.an_other_btn.clicked.connect(win.choose_file)
    btns.addWidget(win.an_other_btn)
    btns.addStretch()
    win.an_run_btn = QPushButton("📊  Анализировать")
    win.an_run_btn.setObjectName("primary")
    win.an_run_btn.setToolTip("Прочитать файлы и создать Excel: организации, что случилось, ошибки")
    win.an_run_btn.clicked.connect(lambda: flow.run(win))
    btns.addWidget(win.an_run_btn)
    lay.addLayout(btns)
    return card
