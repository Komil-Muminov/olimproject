"""Единая таблица глубокого анализа: организации (по ИНН) и их суммы по найденным частям.

Часть — один источник данных: файл реестра (период), лист реестра или таблица с ИНН на листе
обычного файла. Что случилось с организацией, считается по частям, которые выбрал пользователь."""
from dataclasses import dataclass, field

APPEARED, GONE, GREW, FELL, SAME, GAP = "появился", "пропал", "вырос", "упал", "без изменений", "с перерывом"
STATUSES = (APPEARED, GONE, GREW, FELL, SAME, GAP)

DUP_PAY, BAD_INN, NAMES, ODD_DATE, DUP_ROW = "двойной платёж", "неправильный ИНН", "разные названия", \
    "странная дата", "повтор в таблице"
FLAGS = (DUP_PAY, BAD_INN, NAMES, ODD_DATE, DUP_ROW)

PAYER, RECIPIENT, ORG = "плательщик", "получатель", "организация"


@dataclass
class Part:
    label: str          # «Янв–Апр 2024» / «2025 · «Муқоиса» E–G»
    source: str         # имя файла
    kind: str           # "payments" — реестр, "list" — таблица с ИНН
    order: tuple = (9999, 0, 0)      # для порядка «раньше → позже»: (год, месяц, № по порядку)


@dataclass
class Entity:
    inn: str
    name: str
    role: str
    sums: list                       # по частям (0.0, если нет)
    counts: list                     # сколько записей в каждой части (0 — нет в этой части)
    flags: set = field(default_factory=set)

    def total(self, sel: list) -> float:
        return round(sum(self.sums[i] for i in sel), 2)

    def present(self, sel: list) -> bool:
        return any(self.counts[i] for i in sel)


@dataclass
class AnalysisData:
    parts: list                      # [Part] — уже по порядку времени
    entities: list                   # [Entity]
    details: dict = field(default_factory=dict)     # (флаг, "payments"/"list") → [(ключ, строка для Excel)]
    registry: object = None          # свод реестров (для выгрузки платежей одного ИНН)
    paths: list = field(default_factory=list)
    mappings: dict = field(default_factory=dict)

    @property
    def roles(self) -> list:
        found = {e.role for e in self.entities}
        return [r for r in (PAYER, RECIPIENT, ORG) if r in found]


def status(e: Entity, sel: list) -> str:
    """Что случилось между первой и последней выбранной частью."""
    if len(sel) < 2:
        return ""
    first, last = e.counts[sel[0]] > 0, e.counts[sel[-1]] > 0
    if not first and last:
        return APPEARED
    if first and not last:
        return GONE
    if not first and not last:
        return GAP
    a, b = e.sums[sel[0]], e.sums[sel[-1]]
    return GREW if b - a > 0.005 else FELL if a - b > 0.005 else SAME


def entity_key(e: Entity) -> tuple:
    """Ключ для подробностей ошибок: (роль, ИНН как в файле)."""
    return e.role, e.inn
