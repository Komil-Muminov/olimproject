"""Стартовая карточка: ГИБ или Глубокий анализ; для ГИБ — «Заполнить» (2 файла) или «Автоматика» (1 файл)."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup, QComboBox, QFrame, QHBoxLayout, QLabel, QProgressBar,
    QPushButton, QRadioButton, QSizePolicy, QVBoxLayout, QWidget
)

from ui_helpers import BACKUP_NOTE


def _file_row(win, title, attr, slot):
    row = QHBoxLayout()
    cap = QLabel(title)
    cap.setObjectName("fileCaption")
    cap.setFixedWidth(170)
    row.addWidget(cap)
    name = QLabel("не выбран")
    name.setObjectName("muted")
    name.setTextFormat(Qt.PlainText)
    row.addWidget(name, 1)
    btn = QPushButton("Выбрать…")
    btn.clicked.connect(slot)
    row.addWidget(btn)
    setattr(win, attr, name)
    return row


def build_start_card(win):
    card = QFrame()
    card.setObjectName("card")
    lay = QVBoxLayout(card)
    lay.setContentsMargins(28, 24, 28, 24)
    lay.setSpacing(12)

    modes = QHBoxLayout()
    win.mode_group = QButtonGroup(win)
    for i, (text, tip) in enumerate((
            ("🧮   ГИБ", "Распределение сумм по ГИБ из назначений платежей"),
            ("🔍   Глубокий анализ", "Сравнение списков и проверка по ИНН"))):
        b = QPushButton(text)
        b.setObjectName("modeCard")
        b.setCheckable(True)
        b.setToolTip(tip)
        b.setCursor(Qt.PointingHandCursor)
        win.mode_group.addButton(b, i)
        modes.addWidget(b)
    win.mode_group.button(0).setChecked(True)
    win.mode_group.idClicked.connect(lambda _: win.update_start())
    lay.addLayout(modes)

    win.gib_box = QWidget()
    win.gib_box.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
    gl = QVBoxLayout(win.gib_box)
    gl.setContentsMargins(4, 4, 0, 0)
    win.fill_radio = QRadioButton("Заполнить ГИБ — платежи из выписки переносятся в ваш файл и раскладываются по ГИБ")
    win.embed_radio = QRadioButton("Встроить автоматику в Excel — ГИБ будут проставляться сами прямо в Excel")
    win.fill_radio.setChecked(True)
    for rb in (win.fill_radio, win.embed_radio):
        rb.toggled.connect(lambda _: win.update_start())
        gl.addWidget(rb)
    lay.addWidget(win.gib_box)

    win.fill_box = QWidget()
    fl = QVBoxLayout(win.fill_box)
    fl.setContentsMargins(0, 8, 0, 0)
    fl.setSpacing(8)
    fl.addLayout(_file_row(win, "1. Выписка", "st_label", win.choose_statement))
    fl.addLayout(_file_row(win, "2. Куда заполнить", "tg_label", win.choose_target))
    sheet_row = QHBoxLayout()
    win.sheet_caption = QLabel("Лист")
    win.sheet_caption.setObjectName("fileCaption")
    win.sheet_caption.setFixedWidth(170)
    sheet_row.addWidget(win.sheet_caption)
    win.sheet_combo = QComboBox()
    win.sheet_combo.setObjectName("bigSelect")
    win.sheet_combo.currentIndexChanged.connect(lambda _: win.update_fill_info())
    sheet_row.addWidget(win.sheet_combo, 1)
    fl.addLayout(sheet_row)
    win.fill_info = QLabel()
    win.fill_info.setObjectName("bigText")
    win.fill_info.setWordWrap(True)
    win.fill_info.setTextFormat(Qt.RichText)
    fl.addWidget(win.fill_info)
    fl.addStretch()
    note = QLabel(BACKUP_NOTE)
    note.setObjectName("muted")
    note.setWordWrap(True)
    fl.addWidget(note)
    win.fill_progress = QProgressBar()
    win.fill_progress.setVisible(False)
    fl.addWidget(win.fill_progress)
    btns = QHBoxLayout()
    btns.addStretch()
    win.fill_btn = QPushButton("Заполнить ГИБ")
    win.fill_btn.setObjectName("primary")
    win.fill_btn.clicked.connect(win.run_fill)
    btns.addWidget(win.fill_btn)
    fl.addLayout(btns)
    lay.addWidget(win.fill_box, 1)

    win.drop = QPushButton()
    win.drop.setObjectName("dropZone")
    win.drop.clicked.connect(win.choose_file)
    lay.addWidget(win.drop, 1)
    return card
