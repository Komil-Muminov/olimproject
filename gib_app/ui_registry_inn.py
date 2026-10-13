"""Виджет таблицы ТОП-ИНН и автодополнение для карточки реестров платежей."""
import re
from PySide6.QtCore import Qt, QStringListModel
from PySide6.QtWidgets import (
    QButtonGroup, QCompleter, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QRadioButton, QTableWidget, QTableWidgetItem, QVBoxLayout
)

from registry_analysis import party_rows
from utils import format_money


def build_inn_selector(win):
    """Создает компактный блок: переключатель (плательщики/получатели) + таблица ТОП-ИНН."""
    box = QFrame()
    lay = QVBoxLayout(box)
    lay.setContentsMargins(0, 4, 0, 4)
    lay.setSpacing(6)

    top_row = QHBoxLayout()
    lbl = QLabel("Крупнейшие контрагенты (кликните для выбора):")
    lbl.setObjectName("cmpTitle")
    top_row.addWidget(lbl)
    top_row.addStretch()

    win.reg_side_group = QButtonGroup(win)
    win.reg_radio_payers = QRadioButton("Плательщики")
    win.reg_radio_payers.setChecked(True)
    win.reg_radio_recs = QRadioButton("Получатели")
    win.reg_side_group.addButton(win.reg_radio_payers)
    win.reg_side_group.addButton(win.reg_radio_recs)
    top_row.addWidget(win.reg_radio_payers)
    top_row.addWidget(win.reg_radio_recs)
    lay.addLayout(top_row)

    win.reg_table = QTableWidget(0, 4)
    win.reg_table.setObjectName("innTable")
    win.reg_table.setHorizontalHeaderLabels(["ИНН", "Наименование", "Всего сумма", "Платежей"])
    win.reg_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
    win.reg_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
    win.reg_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
    win.reg_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
    win.reg_table.setSelectionBehavior(QTableWidget.SelectRows)
    win.reg_table.setSelectionMode(QTableWidget.SingleSelection)
    win.reg_table.setEditTriggers(QTableWidget.NoEditTriggers)
    win.reg_table.setMinimumHeight(110)
    win.reg_table.setMaximumHeight(160)

    win.reg_table.cellClicked.connect(lambda r, _: select_table_row(win, r))
    win.reg_radio_payers.toggled.connect(lambda: update_inn_table(win))

    lay.addWidget(win.reg_table)
    return box


def select_table_row(win, row: int):
    """Клик по строке таблицы подставляет ИНН и запускает поиск."""
    item = win.reg_table.item(row, 0)
    if not item:
        return
    inn = item.text().strip()
    win.reg_inn.setText(inn)
    import ui_registry_flow as flow
    flow.find(win, inn)


def setup_completer(win, reg):
    """Настраивает автодополнение в строке поиска по всем уникальным ИНН реестра."""
    if not reg:
        win.reg_inn.setCompleter(None)
        return

    items = []
    seen = set()
    for book in (reg.payers, reg.recs):
        for p in party_rows(book, reg):
            if p.inn not in seen:
                seen.add(p.inn)
                items.append(f"{p.inn}  ·  {p.name} ({format_money(p.total)})")

    model = QStringListModel(items, win)
    completer = QCompleter(model, win)
    completer.setCaseSensitivity(Qt.CaseInsensitive)
    completer.setFilterMode(Qt.MatchContains)
    completer.setMaxVisibleItems(10)

    def _on_activated(text):
        inn = text.split()[0] if text else ""
        win.reg_inn.setText(inn)
        import ui_registry_flow as flow
        flow.find(win, inn)

    completer.activated[str].connect(_on_activated)
    win.reg_inn.setCompleter(completer)


def update_inn_table(win, reg=None):
    """Обновляет строки таблицы ТОП-50 по выбранной стороне."""
    table = getattr(win, "reg_table", None)
    reg = reg or getattr(win, "registry", None)
    if not table:
        return
    if not reg:
        table.setRowCount(0)
        return

    is_payers = win.reg_radio_payers.isChecked()
    book = reg.payers if is_payers else reg.recs
    rows = party_rows(book, reg)[:50]

    table.setRowCount(len(rows))
    for r, p in enumerate(rows):
        it_inn = QTableWidgetItem(p.inn)
        it_inn.setTextAlignment(Qt.AlignCenter)
        it_name = QTableWidgetItem(p.name)
        it_name.setToolTip(p.name)
        it_total = QTableWidgetItem(format_money(p.total))
        it_total.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        it_cnt = QTableWidgetItem(str(sum(p.counts)))
        it_cnt.setTextAlignment(Qt.AlignCenter)

        for col, it in enumerate((it_inn, it_name, it_total, it_cnt)):
            table.setItem(r, col, it)


def populate_inn_view(win, reg):
    """Инициализирует таблицу и автодополнение после завершения сканирования."""
    setup_completer(win, reg)
    update_inn_table(win, reg)


def reset_inn_view(win):
    """Сбрасывает состояние таблицы и автодополнения."""
    if hasattr(win, "reg_inn"):
        win.reg_inn.clear()
        win.reg_inn.setCompleter(None)
    if hasattr(win, "reg_table"):
        win.reg_table.setRowCount(0)


def clean_inn_query(raw: str, reg=None) -> str:
    """Извлекает чистый ИНН из введенной строки, подсказки или названия."""
    raw = raw.strip()
    m = re.match(r"^(\d{9,12})", raw)
    if m:
        return m.group(1)
    digits = "".join(c for c in raw if c.isdigit())
    if len(digits) in (9, 10, 12):
        return digits
    if reg and not digits:
        # по названию — самая крупная по сумме из подходящих (остальные видны в подсказках)
        q = raw.lower()
        hits = [p for book in (reg.payers, reg.recs) for p in book.values() if q in p.name.lower()]
        if hits:
            return max(hits, key=lambda p: p.total).inn
    return digits or raw
