"""Список ИНН на карточке реестров: крупнейшие сверху, при вводе в поиск — фильтр по ИНН или названию.
Фильтр запускается через 250 мс после последней буквы, поэтому окно не подвисает на 90 тыс. ИНН."""
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView, QButtonGroup, QFrame, QHBoxLayout, QHeaderView, QLabel, QRadioButton,
    QTableWidget, QTableWidgetItem, QVBoxLayout
)

from registry_index import build_index, matching
from utils import format_money

SHOWN = 50
DELAY_MS = 250
AUTO_COLS = (0, 2, 3)        # ширина по содержимому; «Название» растягивается


def build_inn_selector(win):
    """Заголовок списка + «Плательщики / Получатели» + таблица (клик — открыть ИНН)."""
    box = QFrame()
    lay = QVBoxLayout(box)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(6)
    top = QHBoxLayout()
    win.reg_list_title = QLabel("")
    win.reg_list_title.setObjectName("muted")
    top.addWidget(win.reg_list_title)
    top.addStretch()
    win.reg_side_group = QButtonGroup(win)
    win.reg_radio_payers = QRadioButton("Плательщики")
    win.reg_radio_payers.setChecked(True)
    win.reg_radio_recs = QRadioButton("Получатели")
    for rb in (win.reg_radio_payers, win.reg_radio_recs):
        win.reg_side_group.addButton(rb)
        top.addWidget(rb)
    lay.addLayout(top)

    t = win.reg_table = QTableWidget(0, 4)
    t.setObjectName("innTable")
    t.setHorizontalHeaderLabels(["ИНН", "Название", "Всего", "Платежей"])
    t.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
    _auto_width(t, True)
    t.verticalHeader().setVisible(False)
    t.verticalHeader().setDefaultSectionSize(30)
    t.setShowGrid(False)
    t.setAlternatingRowColors(True)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setSelectionMode(QAbstractItemView.SingleSelection)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setCursor(Qt.PointingHandCursor)
    t.cellClicked.connect(lambda r, _: select_table_row(win, r))
    lay.addWidget(t, 1)

    win.reg_filter_timer = QTimer(win)
    win.reg_filter_timer.setSingleShot(True)
    win.reg_filter_timer.setInterval(DELAY_MS)
    win.reg_filter_timer.timeout.connect(lambda: update_inn_table(win))
    win.reg_radio_payers.toggled.connect(lambda: update_inn_table(win))
    return box


def _auto_width(t, on: bool):
    """Автоширину выключаем на время заполнения: иначе таблица пересчитывает её после каждой ячейки."""
    for col in AUTO_COLS:
        t.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents if on else QHeaderView.Fixed)


def on_text_edited(win):
    """Ввод в поиск: вернуть список и отфильтровать его чуть позже (не на каждую букву)."""
    win.reg_view.setCurrentIndex(0)
    win.reg_filter_timer.start()


def select_table_row(win, row: int):
    item = win.reg_table.item(row, 0)
    if item:
        import ui_registry_flow as flow
        flow.find(win, item.text().strip())


def update_inn_table(win, reg=None):
    """Строки таблицы: крупнейшие (поиск пуст) или подходящие под текст поиска."""
    reg = reg or getattr(win, "registry", None)
    index = getattr(win, "reg_index", None)
    if not reg or not index:
        win.reg_table.setRowCount(0)
        win.reg_list_title.setText("")
        return
    side = "payer" if win.reg_radio_payers.isChecked() else "rec"
    who = "плательщиков" if side == "payer" else "получателей"
    q = win.reg_inn.text().strip().lower()
    rows = matching(index[side], q)
    if q:
        win.reg_list_title.setText(f"Найдено {who}: {len(rows):,}".replace(",", " ")
                                   + (f" — показаны {SHOWN} крупнейших" if len(rows) > SHOWN else "")
                                   + " · нажмите на строку")
    else:
        win.reg_list_title.setText(f"Крупнейшие {SHOWN} {who} из {len(rows):,}".replace(",", " ")
                                   + " · нажмите на строку или начните вводить ИНН / название")
    rows = rows[:SHOWN]
    t = win.reg_table
    t.setUpdatesEnabled(False)          # без перерисовки и пересчёта ширины на каждую ячейку
    _auto_width(t, False)
    t.setRowCount(len(rows))
    for i, r in enumerate(rows):
        cells = ((r.inn, Qt.AlignCenter), (r.name, Qt.AlignLeft | Qt.AlignVCenter),
                 (format_money(r.total), Qt.AlignRight | Qt.AlignVCenter), (str(sum(r.counts)), Qt.AlignCenter))
        for col, (text, align) in enumerate(cells):
            it = QTableWidgetItem(text)
            it.setTextAlignment(align)
            if col == 1:
                it.setToolTip(text)
            t.setItem(i, col, it)
    _auto_width(t, True)
    t.setUpdatesEnabled(True)


def populate_inn_view(win, reg):
    """После чтения реестров: индекс для фильтра и таблица крупнейших."""
    win.reg_index = build_index(reg)
    update_inn_table(win, reg)


def reset_inn_view(win):
    win.reg_index = None
    win.reg_inn.clear()
    win.reg_table.setRowCount(0)
    win.reg_list_title.setText("")
    win.reg_view.setCurrentIndex(0)
