"""Глубокий анализ: организации (по ИНН) и их суммы по найденным частям.

Часть — один источник данных: файл реестра (период), лист реестра или таблица с ИНН на листе
обычного файла. Что случилось с организацией, считается между первой и последней частью."""
from dataclasses import dataclass, field

APPEARED, GONE, GREW, FELL, SAME, GAP = "Новый", "Пропал", "Вырос", "Уменьшился", "Без изменений", \
    "Появился и пропал"
STATUSES = (APPEARED, GONE, GREW, FELL, SAME, GAP)
STATUS_HINTS = {
    APPEARED: "В первой части этой организации нет, в последней — есть.",
    GONE: "В первой части организация есть, в последней — нет.",
    GREW: "Есть и в первой, и в последней части; в последней сумма больше.",
    FELL: "Есть и в первой, и в последней части; в последней сумма меньше.",
    SAME: "Есть и в первой, и в последней части; сумма одинаковая.",
    GAP: "Нет ни в первой, ни в последней части — была только в средних.",
}

DUP_PAY, BAD_INN, ODD_DATE, DUP_ROW = "двойной платёж", "неправильный ИНН", "странная дата", "повтор ИНН"
FLAGS = (DUP_PAY, BAD_INN, ODD_DATE, DUP_ROW)
FLAG_HINTS = {
    DUP_PAY: "Один и тот же платёж записан два раза или больше: одинаковые №, дата, ИНН плательщика, "
             "ИНН получателя и сумма.",
    BAD_INN: "ИНН не из 9 цифр: «0», пустой, с буквами или потерян ноль в начале.",
    ODD_DATE: "Дата платёжки не из периода файла (например, 2004 или 2029 год) — возможно, опечатка.",
    DUP_ROW: "Один ИНН записан в одной таблице несколько раз.",
}

PAYER, RECIPIENT, ORG = "плательщик", "получатель", "организация"
DETAIL_HEAD = ["Ошибка", "Файл", "Где", "ИНН", "Название", "№ док.", "Дата", "Сумма", "Пояснение"]


@dataclass
class Part:
    label: str          # «Янв–Апр 2024» / «2025 · «Муқоиса» E–G»
    source: str         # имя файла
    kind: str           # "payments" — реестр, "list" — таблица с ИНН
    order: tuple = (9999, 0, 0)      # для порядка «раньше → позже»: (год, месяц, № по порядку)
    has_sums: bool = True            # в таблице есть столбец сумм (иначе — только «есть / нет»)


@dataclass
class Entity:
    inn: str
    name: str
    role: str
    sums: list                       # по частям (0.0, если нет)
    counts: list                     # сколько записей в каждой части (0 — нет в этой части)
    flags: set = field(default_factory=set)
    names: int = 1                   # сколько разных написаний названия

    @property
    def total(self) -> float:
        return round(sum(self.sums), 2)


@dataclass
class AnalysisData:
    parts: list                      # [Part] — по порядку времени
    entities: list                   # [Entity]
    details: list = field(default_factory=list)     # строки листа «Ошибки подробно» (DETAIL_HEAD)
    paths: list = field(default_factory=list)

    @property
    def roles(self) -> list:
        found = {e.role for e in self.entities}
        return [r for r in (PAYER, RECIPIENT, ORG) if r in found]


def status(e: Entity, parts: list) -> str:
    """Что случилось между первой и последней частью ("" — если часть одна). Если у одной из них нет
    сумм, сравнивается только наличие: есть в обеих — «Без изменений»."""
    if len(e.counts) < 2:
        return ""
    first, last = e.counts[0] > 0, e.counts[-1] > 0
    if not first and last:
        return APPEARED
    if first and not last:
        return GONE
    if not first and not last:
        return GAP
    if not (parts[0].has_sums and parts[-1].has_sums):
        return SAME
    a, b = e.sums[0], e.sums[-1]
    return GREW if b - a > 0.005 else FELL if a - b > 0.005 else SAME
