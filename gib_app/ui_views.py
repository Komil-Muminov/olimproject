"""Компоновка карточек интерфейса (карточка анализа и карточка результата)."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup, QFrame, QHBoxLayout, QLabel,
    QProgressBar, QPushButton, QVBoxLayout
)


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

    win.cmp_row = QFrame()
    win.cmp_row.setObjectName("cmpBox")
    cl = QVBoxLayout(win.cmp_row)
    cl.setContentsMargins(0, 4, 0, 0)
    cl.setSpacing(6)

    title = QLabel("Сравнение по ИНН  ·  выберите лист")
    title.setObjectName("cmpTitle")
    cl.addWidget(title)

    win.sheet_opts = QVBoxLayout()
    win.sheet_opts.setSpacing(2)
    cl.addLayout(win.sheet_opts)

    win.sheet_group = QButtonGroup(win)
    win.sheet_group.setExclusive(True)
    win.sheet_group.idClicked.connect(win._compare_sheet_changed)

    win.cmp_info = QLabel()
    win.cmp_info.setObjectName("cmpInfo")
    win.cmp_info.setWordWrap(True)
    win.cmp_info.setTextFormat(Qt.RichText)
    cl.addWidget(win.cmp_info)
    fl.addWidget(win.cmp_row)

    win.warn_text = QLabel()
    win.warn_text.setObjectName("warn")
    win.warn_text.setWordWrap(True)
    win.warn_text.setTextFormat(Qt.RichText)
    fl.addWidget(win.warn_text)
    fl.addStretch()

    win.note = QLabel("🛡  Перед изменением программа сохранит копию файла в папку «_резервные копии» (в папке программы).")
    win.note.setObjectName("muted")
    win.note.setWordWrap(True)
    fl.addWidget(win.note)

    win.progress = QProgressBar()
    win.progress.setVisible(False)
    fl.addWidget(win.progress)

    btns = QHBoxLayout()
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

    win.cmp_btn = QPushButton("🔍  Сравнить по ИНН")
    win.cmp_btn.setObjectName("primary")
    win.cmp_btn.setToolTip("Создаёт новый файл с различиями двух списков: новые, пропавшие,\n"
                            "изменившиеся суммы и дубликаты. Исходный файл не меняется.")
    win.cmp_btn.clicked.connect(win.run_compare)
    btns.addWidget(win.cmp_btn)

    win.chk_btn = QPushButton("🧾  Проверить ИНН")
    win.chk_btn.setObjectName("secondary")
    win.chk_btn.setToolTip("Создаёт новый файл: дубликаты ИНН, один ИНН — разные названия,\n"
                            "неправильные ИНН. Проверяются все таблицы с ИНН на всех листах.")
    win.chk_btn.clicked.connect(win.run_check)
    btns.addWidget(win.chk_btn)

    win.go_btn = QPushButton("Заполнить ГИБ")
    win.go_btn.setObjectName("primary")
    win.go_btn.clicked.connect(win.run)
    btns.addWidget(win.go_btn)

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
    again = QPushButton("Открыть другой файл")
    again.clicked.connect(win.choose_file)
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
