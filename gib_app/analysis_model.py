"""Глубокий анализ: организации (по ИНН) по найденным частям.

Часть — один источник данных: таблица с ИНН на листе файла. Сравнения «было → стало» нет: по каждой
организации — на каких листах она есть и в каких файлах."""
from dataclasses import dataclass, field

BAD_INN, DUP_ROW = "неправильный ИНН", "повтор ИНН"
SAME_NAME = "одно название — разные ИНН"
DIFF_NAMES = "разные названия"
SHARED, BY_NAME, HIDDEN = "общая строка", "по названию", "скрытая строка"   # пометки (не ошибки): см. FLAG_HINTS
FLAGS = (BAD_INN, DUP_ROW, SAME_NAME, DIFF_NAMES, SHARED, BY_NAME, HIDDEN)
FLAG_HINTS = {
    BAD_INN: "ИНН не из 9 цифр: «0», пустой, с буквами или потерян ноль в начале.",
    DUP_ROW: "Один ИНН записан в одной таблице несколько раз.",
    SAME_NAME: "У организации с таким же названием есть другой ИНН — возможно, ошибка в ИНН (или это филиалы).",
    DIFF_NAMES: "У этого ИНН в файлах записаны разные названия — наведите на название, чтобы увидеть все.",
    SHARED: "В строке списка было несколько организаций («ООО А, ООО Б») — сумма строки записана каждой целиком.",
    BY_NAME: "В списке нет ИНН — организация сопоставлена по названию (менее надёжно: проверьте).",
    HIDDEN: "Есть в скрытой строке файла (в Excel её не видно) — проверьте, почему строку скрыли.",
}

ORG = "организация"


@dataclass
class Part:
    label: str          # «Янв–Апр 2024» / «2025 · «Муқоиса» E–G»
    source: str         # имя файла
    kind: str           # "list" — таблица с ИНН
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
class AnalysisData:
    parts: list                      # [Part] — по порядку времени
    entities: list                   # [Entity]
    details: dict = field(default_factory=dict)     # группа ошибок → строки (report_errors.GROUPS)
    paths: list = field(default_factory=list)
    notes: list = field(default_factory=list)        # скрытые листы и строки (excel_sheet_info.hidden_note)
    checks: list = field(default_factory=list)       # сверка с «Итого» в файле (excel_sheet_info.total_check)


def present_files(e, parts: list) -> list:
    """Файлы, где есть организация (по порядку частей, без повторов: листы одного файла — один файл)."""
    return list(dict.fromkeys(p.source for i, p in enumerate(parts) if e.counts[i]))


def part_summary(data: AnalysisData) -> list:
    """По каждой части: организаций и сумма (если в части есть суммы)."""
    out = []
    for i, p in enumerate(data.parts):
        here = [e for e in data.entities if e.counts[i]]
        out.append({
            "part": p, "orgs": len(here),
            "sum": round(sum(e.sums[i] for e in here), 2) if p.has_sums else None,
        })
    return out
