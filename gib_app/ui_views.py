"""Шапка окна, карточка файла (для «Встроить автоматику») и карточка результата."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel,
    QProgressBar, QPushButton, QVBoxLayout
)

import ui_analysis
import ui_analysis_flow
from ui_helpers import BACKUP_NOTE
from ui_result import chips_box


def build_header(win):
    """Название программы, подпись и кнопка «Справочник ГИБ»."""
    header = QHBoxLayout()
    info = QVBoxLayout()
    title = QLabel("Распределение по ГИБ")
    title.setObjectName("appTitle")
    info.addWidget(title)
    sub = QLabel("ГИБ — распределение сумм по назначениям. Глубокий анализ — Excel по организациям (ИНН): "
                 "что случилось и ошибки.")
    sub.setObjectName("muted")
    info.addWidget(sub)
    header.addLayout(info)
    header.addStretch()
    win.catalog_btn = QPushButton("📖  Справочник ГИБ")
    win.catalog_btn.setObjectName("secondary")
    win.catalog_btn.setToolTip("Просмотр и редактирование статей расходов бюджетной классификации")
    win.catalog_btn.clicked.connect(win.open_catalog)
    header.addWidget(win.catalog_btn)
    return header


def build_found_card(win):
    card = QFrame()
    card.setObjectName("card")
    fl = QVBoxLayout(card)
    fl.setContentsMargins(28, 24, 28, 24)
    fl.setSpacing(12)

    win.file_label = QLabel()
    win.file_label.setObjectName("fileName")
    fl.addWidget(win.file_label)

    win.found_text = QLabel()
    win.found_text.setObjectName("bigText")
    win.found_text.setWordWrap(True)
    win.found_text.setTextFormat(Qt.RichText)
    fl.addWidget(win.found_text)

    win.warn_text = QLabel()
    win.warn_text.setObjectName("warn")
    win.warn_text.setWordWrap(True)
    win.warn_text.setTextFormat(Qt.RichText)
    fl.addWidget(win.warn_text)
    fl.addStretch()

    win.note = QLabel(BACKUP_NOTE)
    win.note.setObjectName("muted")
    win.note.setWordWrap(True)
    fl.addWidget(win.note)

    win.progress = QProgressBar()
    win.progress.setVisible(False)
    fl.addWidget(win.progress)

    btns = QHBoxLayout()
    back = QPushButton("←  Назад")
    back.clicked.connect(win.go_start)
    btns.addWidget(back)
    other = QPushButton("Другой файл")
    other.clicked.connect(win.choose_file)
    btns.addWidget(other)

    view = QPushButton("Открыть в Excel")
    view.clicked.connect(win.open_file)
    btns.addWidget(view)
    btns.addStretch()

    win.embed_btn = QPushButton("⚙  Встроить автоматику в Excel")
    win.embed_btn.setObjectName("secondary")
    win.embed_btn.setToolTip("Создаёт копию файла с макросом: дальше ГИБ проставляются сами,\n"
                              "когда вы вставляете назначения прямо в Excel.")
    win.embed_btn.clicked.connect(win.embed)
    btns.addWidget(win.embed_btn)

    fl.addLayout(btns)
    return card


def build_result_card(win):
    card = QFrame()
    card.setObjectName("card")
    rl = QVBoxLayout(card)
    rl.setContentsMargins(28, 24, 28, 24)
    rl.setSpacing(12)

    win.result_head = QLabel()
    win.result_head.setObjectName("resultOk")
    rl.addWidget(win.result_head)

    win.result_text = QLabel()
    win.result_text.setObjectName("bigText")
    win.result_text.setWordWrap(True)
    win.result_text.setTextFormat(Qt.RichText)
    rl.addWidget(win.result_text)
    win.result_chips = chips_box()
    rl.addWidget(win.result_chips)
    win.result_after = QLabel()             # текст под метками (например, ошибки по файлам)
    win.result_after.setObjectName("bigText")
    win.result_after.setWordWrap(True)
    win.result_after.setTextFormat(Qt.RichText)
    rl.addWidget(win.result_after)
    # поле ИНН — показывается только после глубокого анализа
    win.result_inn_box, win.result_inn, _ = ui_analysis.inn_box(lambda: ui_analysis_flow.inn_from_result(win))
    win.result_inn_box.setVisible(False)
    rl.addWidget(win.result_inn_box)

    win.result_warn = QLabel()
    win.result_warn.setObjectName("warn")
    win.result_warn.setWordWrap(True)
    win.result_warn.setTextFormat(Qt.RichText)
    rl.addWidget(win.result_warn)
    rl.addStretch()

    win.backup_label = QLabel()
    win.backup_label.setObjectName("muted")
    win.backup_label.setWordWrap(True)
    rl.addWidget(win.backup_label)

    rb = QHBoxLayout()
    again = QPushButton("←  В начало")
    again.clicked.connect(win.go_start)
    rb.addWidget(again)
    rb.addStretch()

    win.backup_btn = QPushButton("Открыть копию (до изменений)")
    win.backup_btn.clicked.connect(lambda: win._start_file(win.backup))
    rb.addWidget(win.backup_btn)

    win.open_btn = QPushButton("Открыть в Excel")
    win.open_btn.setObjectName("primary")
    win.open_btn.clicked.connect(lambda: win._start_file(win.result_path))
    rb.addWidget(win.open_btn)

    rl.addLayout(rb)
    return card
