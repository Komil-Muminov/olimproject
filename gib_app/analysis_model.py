"""Глубокий анализ: организации (по ИНН) и их суммы по найденным частям.

Часть — один источник данных: файл реестра (период), лист реестра или таблица с ИНН на листе
обычного файла. Что случилось с организацией, считается между первой и последней частью."""
from dataclasses import dataclass, field

APPEARED, GONE, GREW, FELL, SAME, GAP = "Новый", "Пропал", "Вырос", "Уменьшился", "Без изменений", \
    "Появился и пропал"
STATUSES = (APPEARED, GONE, GREW, FELL, SAME, GAP)


def status_hints(parts: list) -> dict:
    """Подсказки к цветам — с названиями первой и последней части: понятно, что с чем сравнивается."""
    if len(parts) < 2:
        return {}
    a, b = f"«{parts[0].label}»", f"«{parts[-1].label}»"
    mid = ", ".join(f"«{p.label}»" for p in parts[1:-1])
    return {
        APPEARED: f"Новый: в {a} этой организации не было, а в {b} она есть.",
        GONE: f"Пропал: в {a} организация была, а в {b} её уже нет.",
        GREW: f"Вырос: организация есть и в {a}, и в {b}; в {b} сумма больше, чем в {a}.",
        FELL: f"Уменьшился: организация есть и в {a}, и в {b}; в {b} сумма меньше, чем в {a}.",
        SAME: f"Без изменений: организация есть и в {a}, и в {b}, сумма одинаковая "
              f"(если в одной из таблиц сумм нет — значит, просто есть и там, и там).",
        GAP: f"Появился и пропал: нет ни в {a}, ни в {b} — была только в середине"
             + (f" ({mid})." if mid else "."),
    }

DUP_PAY, BAD_INN, ODD_DATE, DUP_ROW = "двойной платёж", "неправильный ИНН", "странная дата", "повтор ИНН"
FLAGS = (DUP_PAY, BAD_INN, ODD_DATE, DUP_ROW)
FLAG_HINTS = {
    DUP_PAY: "Один и тот же платёж записан два раза или больше: одинаковые №, дата, ИНН плательщика, "
             "ИНН получателя и сумма.",
    BAD_INN: "ИНН не из 9 цифр: «0», пустой, с буквами или потерян ноль в начале.",
    ODD_DATE: "Дата платёжки больше чем на месяц раньше или позже периода файла (например, 2004 или 2029 год) — "
              "возможно, опечатка в дате.",
    DUP_ROW: "Один ИНН записан в одной таблице несколько раз.",
}

PAYER, RECIPIENT, ORG = "плательщик", "получатель", "организация"


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
    details: dict = field(default_factory=dict)     # группа ошибок → строки (report_errors.GROUPS)
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


def part_summary(data: AnalysisData) -> list:
    """По каждой части: организаций, сумма, сколько появилось и пропало по сравнению с предыдущей частью.
    Сумма платежей — по плательщикам (иначе каждый платёж посчитается дважды)."""
    money_role = PAYER if PAYER in data.roles else None
    out = []
    for i, p in enumerate(data.parts):
        here = [e for e in data.entities if e.counts[i]]
        prev = [e.counts[i - 1] > 0 for e in data.entities] if i else None
        out.append({
            "part": p, "orgs": len(here),
            "sum": round(sum(e.sums[i] for e in here if money_role is None or e.role == money_role), 2)
            if p.has_sums else None,
            "new": sum(1 for k, e in enumerate(data.entities) if e.counts[i] and not prev[k]) if i else None,
            "gone": sum(1 for k, e in enumerate(data.entities) if prev[k] and not e.counts[i]) if i else None,
        })
    return out
