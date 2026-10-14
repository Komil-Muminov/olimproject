"""Экран глубокого анализа: выбранные файлы и два действия — «Анализ» и платежи по ИНН."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QLineEdit, QProgressBar, QPushButton, QSizePolicy, QVBoxLayout
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

    inn_box = QFrame()
    inn_box.setObjectName("actionBox")
    inn_lay = QVBoxLayout(inn_box)
    inn_lay.setContentsMargins(18, 14, 18, 14)
    inn_lay.setSpacing(8)
    title = QLabel("📄  Всё по одной организации")
    title.setObjectName("actionTitle")
    inn_lay.addWidget(title)
    inn_row = QHBoxLayout()
    win.an_inn = QLineEdit()
    win.an_inn.setPlaceholderText("ИНН, например 040000011")
    win.an_inn.setMaxLength(20)
    win.an_inn.setClearButtonEnabled(True)
    win.an_inn.returnPressed.connect(lambda: flow.export_payments(win))
    inn_row.addWidget(win.an_inn, 1)
    win.an_pay_btn = QPushButton("В Excel")
    win.an_pay_btn.setToolTip("Где встречается этот ИНН во всех выбранных файлах: строки таблиц (с подписью — "
                              "файл, лист, столбец) и платежи из реестров (где платила и где получала)")
    win.an_pay_btn.clicked.connect(lambda: flow.export_payments(win))
    inn_row.addWidget(win.an_pay_btn)
    inn_lay.addLayout(inn_row)
    lay.addWidget(inn_box)
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
