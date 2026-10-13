"""Модели данных и структуры для сравнения списков организаций."""
import re
from dataclasses import dataclass, field
from openpyxl.utils import get_column_letter

from utils import norm_text


def metric_key(h) -> str:
    """Показатель без годов и дат: «Бакия ба 01.01.2024 (30 фоиза)» → «бакия ба (30 фоиза)»."""
    t = norm_text(h)
    t = re.sub(r"\d{1,2}[./]\d{1,2}[./]\d{2,4}", " ", t)
    t = re.sub(r"(?<!\d)20\d\d(?!\d)", " ", t)
    return " ".join(t.split())


def metric_title(h) -> str:
    """Название показателя без даты и года: «Бакия ба 01.01.2024 (30 фоиза)» → «Бакия (30 фоиза)»."""
    t = str(h or "").replace("\n", " ")
    t = re.sub(r"\s*ба\s+\d{1,2}[./]\d{1,2}[./]\d{2,4}", "", t)
    t = re.sub(r"\d{1,2}[./]\d{1,2}[./]\d{2,4}|(?<!\d)20\d\d(?!\d)", "", t)
    return " ".join(t.split())


@dataclass
class SideList:
    key_col: int
    name_col: int | None
    cols: list                                   # столбцы списка (для подписи)
    value_cols: list                             # числовые столбцы
    rows: dict = field(default_factory=dict)     # ключ сравнения -> строка
    keys: dict = field(default_factory=dict)     # ключ сравнения -> ИНН как в файле
    all_rows: dict = field(default_factory=dict) # ключ сравнения -> все строки (для дубликатов)
    year: str | None = None
    block: list = field(default_factory=list)    # столбцы самого списка (№, название, ИНН…)
    side: str = ""                               # «левая часть» / «правая часть» листа
    header: int = 0                              # строка шапки
    last_row: int = 0                            # последняя заполненная строка таблицы

    @property
    def span(self) -> str:
        """«A–C» — где список виден в Excel."""
        return f"{get_column_letter(min(self.block))}–{get_column_letter(max(self.block))}"

    def where(self) -> str:
        """«левая часть — столбцы A–C (2024, 381 орг.)»"""
        year = f"{self.year}, " if self.year else ""
        return f"{self.side} ({self.span}, {year}{len(self.rows)} орг.)"

    @property
    def short(self) -> str:
        """Короткое имя для подписей: год или «правой части (E–G)»."""
        return self.year or f"{self.side} ({self.span})"

    def label(self, i: int) -> str:
        span = f"{get_column_letter(min(self.cols))}–{get_column_letter(max(self.cols))}"
        year = f" · {self.year}" if self.year else ""
        return f"Список {i}{year} — {len(self.rows)} орг. (столбцы {span})"


@dataclass
class Diff:
    only1: list          # (ИНН, название, {показатель: значение})
    only2: list
    changed: list        # (ИНН, название, показатель, было, стало, разница)
    same: int            # организаций без изменений
    common: int
    names_differ: int
    pairs: list          # [(показатель, столбец1, столбец2)]
    duplicates: list = field(default_factory=list)   # (номер списка, ИНН, название, [строки])
    totals: list = field(default_factory=list)       # (показатель, итог списка 1, итог списка 2)
    org_only1: list = field(default_factory=list)
    org_only2: list = field(default_factory=list)
    org_changed: list = field(default_factory=list)
    org_same: list = field(default_factory=list)
    org_names: list = field(default_factory=list)    # (ИНН, название 1, название 2)
