"""Глубокий анализ: организации (по ИНН) и их суммы по найденным частям.

Часть — один источник данных: файл реестра (период), лист реестра или таблица с ИНН на листе
обычного файла. Сравнения «было → стало» нет: по каждой организации — суммы по частям и в каких файлах она есть."""
from dataclasses import dataclass, field

DUP_PAY, BAD_INN, ODD_DATE, DUP_ROW = "двойной платёж", "неправильный ИНН", "странная дата", "повтор ИНН"
SAME_NAME = "одно название — разные ИНН"
DIFF_NAMES = "разные названия"
SHARED, BY_NAME, HIDDEN = "общая строка", "по названию", "скрытая строка"   # пометки (не ошибки): см. FLAG_HINTS
FLAGS = (DUP_PAY, BAD_INN, ODD_DATE, DUP_ROW, SAME_NAME, DIFF_NAMES, SHARED, BY_NAME, HIDDEN)
FLAG_HINTS = {
    DUP_PAY: "Один и тот же платёж записан два раза или больше: одинаковые №, дата, ИНН плательщика, "
             "ИНН получателя и сумма.",
    BAD_INN: "ИНН не из 9 цифр: «0», пустой, с буквами или потерян ноль в начале.",
    ODD_DATE: "Дата платёжки больше чем на месяц раньше или позже периода файла (например, 2004 или 2029 год) — "
              "возможно, опечатка в дате.",
    DUP_ROW: "Один ИНН записан в одной таблице несколько раз.",
    SAME_NAME: "У организации с таким же названием есть другой ИНН — возможно, ошибка в ИНН (или это филиалы).",
    DIFF_NAMES: "У этого ИНН в файлах записаны разные названия — наведите на название, чтобы увидеть все.",
    SHARED: "В строке списка было несколько организаций («ООО А, ООО Б») — сумма строки записана каждой целиком.",
    BY_NAME: "В списке нет ИНН — организация сопоставлена по названию (менее надёжно: проверьте).",
    HIDDEN: "Есть в скрытой строке файла (в Excel её не видно) — проверьте, почему строку скрыли.",
}

PAYER, RECIPIENT, ORG = "плательщик", "получатель", "организация"


@dataclass
class Part:
    label: str          # «Янв–Апр 2024» / «2025 · «Муқоиса» E–G»
    source: str         # имя файла
    kind: str           # "payments" — реестр, "list" — таблица с ИНН
    order: tuple = (9999, 0, 0)      # для порядка «раньше → позже»: (год, месяц, № по порядку)
    has_sums: bool = True            # в таблице есть столбец сумм (иначе — только «есть / нет»)
    heads: list = field(default_factory=list)   # списки: заголовки столбцов сумм (отмеченных пользователем)


@dataclass
class Entity:
    inn: str
    name: str
    role: str
    sums: list                       # по частям (0.0, если нет)
    counts: list                     # сколько записей в каждой части (0 — нет в этой части)
    flags: set = field(default_factory=set)
    names: int = 1                   # сколько разных написаний названия
    cols: list = field(default_factory=list)    # по частям: значения столбцов сумм списка по отдельности (или None)
    variants: list = field(default_factory=list)  # разные написания названия (самое частое первым)

    @property
    def total(self) -> float:
        return round(sum(self.sums), 2)


@dataclass
class Pair:
    """Кто кому платит: плательщик → получатель, суммы и платежи по частям."""
    payer_inn: str
    payer: str
    rec_inn: str
    rec: str
    sums: list
    counts: list


@dataclass
class AnalysisData:
    parts: list                      # [Part] — по порядку времени
    entities: list                   # [Entity]
    details: dict = field(default_factory=dict)     # группа ошибок → строки (report_errors.GROUPS)
    paths: list = field(default_factory=list)
    pairs: list = field(default_factory=list)        # [Pair] — только для реестров
    mode: str = "files"                              # files / quarters / months
    months: dict = field(default_factory=dict)       # реестры: (год, месяц) → [сумма, платежей] — для графика
    notes: list = field(default_factory=list)        # скрытые листы и строки (excel_sheet_info.hidden_note)
    checks: list = field(default_factory=list)       # сверка с «Итого» в файле (excel_sheet_info.total_check)

    @property
    def roles(self) -> list:
        found = {e.role for e in self.entities}
        return [r for r in (PAYER, RECIPIENT, ORG) if r in found]


def present_files(e, parts: list) -> list:
    """Файлы, где есть организация (по порядку частей, без повторов: листы одного файла — один файл)."""
    return list(dict.fromkeys(p.source for i, p in enumerate(parts) if e.counts[i]))


def part_summary(data: AnalysisData) -> list:
    """По каждой части: организаций и сумма. Сумма платежей — по плательщикам (иначе каждый платёж посчитается дважды)."""
    money_role = PAYER if PAYER in data.roles else None
    out = []
    for i, p in enumerate(data.parts):
        here = [e for e in data.entities if e.counts[i]]
        out.append({
            "part": p, "orgs": len(here),
            "sum": round(sum(e.sums[i] for e in here if money_role is None or e.role == money_role), 2)
            if p.has_sums else None,
        })
    return out
