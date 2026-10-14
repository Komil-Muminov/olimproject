"""Модель таблицы глубокого анализа для QTableView: быстро на 90 тыс. строк, сортировка по столбцу."""
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt

from utils import format_money

RIGHT = Qt.AlignRight | Qt.AlignVCenter


class AnalysisTable(QAbstractTableModel):
    def __init__(self):
        super().__init__()
        self.rows, self.labels, self.sel, self.with_role = [], [], [], False

    def set_rows(self, rows: list, labels: list, sel: list, with_role: bool):
        """rows — [(организация, что случилось, всего)]; labels — подписи выбранных частей."""
        self.beginResetModel()
        self.rows, self.labels, self.sel, self.with_role = rows, labels, sel, with_role
        self.endResetModel()

    def _heads(self) -> list:
        return ["ИНН", "Название"] + (["Кто"] if self.with_role else []) + self.labels + \
            ["Всего", "Что случилось", "Ошибки"]

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._heads())

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return self._heads()[section]
        return None

    def _value(self, row, col):
        """(значение для сортировки, текст)."""
        e, st, total = row
        base = 3 if self.with_role else 2
        if col == 0:
            return e.inn, e.inn
        if col == 1:
            return e.name.lower(), e.name
        if self.with_role and col == 2:
            return e.role, e.role
        if base <= col < base + len(self.sel):
            i = self.sel[col - base]
            return (e.sums[i], format_money(e.sums[i])) if e.counts[i] else (float("-inf"), "—")
        k = col - base - len(self.sel)
        if k == 0:
            return total, format_money(total)
        if k == 1:
            return st, st
        flags = ", ".join(sorted(e.flags))
        return flags, flags

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row, col = self.rows[index.row()], index.column()
        if role in (Qt.DisplayRole, Qt.ToolTipRole):
            return self._value(row, col)[1]
        if role == Qt.TextAlignmentRole:
            base = 3 if self.with_role else 2
            return RIGHT if base <= col <= base + len(self.sel) else None
        return None

    def sort(self, column, order=Qt.AscendingOrder):
        self.layoutAboutToBeChanged.emit()
        self.rows.sort(key=lambda r: self._value(r, column)[0], reverse=order == Qt.DescendingOrder)
        self.layoutChanged.emit()

    def entity_at(self, row: int):
        return self.rows[row][0] if 0 <= row < len(self.rows) else None
