"""Фильтры таблицы анализа: части, кто, что случилось, ошибки, сумма, поиск.
Счётчики у вариантов «что случилось» и «ошибки» считаются по строкам после остальных фильтров."""
from collections import Counter
from dataclasses import dataclass, field

from analysis_model import AnalysisData, status


@dataclass
class Filters:
    parts: list                                  # выбранные части (номера по порядку времени)
    role: str | None = None                      # None — все
    statuses: set = field(default_factory=set)   # пусто — любые
    flags: set = field(default_factory=set)      # пусто — с ошибками и без
    min_sum: float | None = None
    max_sum: float | None = None
    query: str = ""


def _match(e, q: str) -> bool:
    if q.isdigit():
        return q.lstrip("0") in e.inn.lstrip("0")
    return q in e.name.lower()


def apply(data: AnalysisData, f: Filters) -> tuple:
    """→ (строки [(организация, что случилось, всего)], счётчики «что случилось», счётчики ошибок).
    Строки — крупные сверху."""
    sel, q = sorted(f.parts), " ".join(f.query.lower().split())
    base = []
    for e in data.entities:
        if not e.present(sel) or (f.role and e.role != f.role) or (q and not _match(e, q)):
            continue
        total = e.total(sel)
        if (f.min_sum is not None and total < f.min_sum) or (f.max_sum is not None and total > f.max_sum):
            continue
        base.append((e, status(e, sel), total))
    st_counts = Counter(st for _, st, _ in base if st)
    flag_counts = Counter(flag for e, _, _ in base for flag in e.flags)
    rows = [r for r in base if (not f.statuses or r[1] in f.statuses) and (not f.flags or r[0].flags & f.flags)]
    rows.sort(key=lambda r: -r[2])
    return rows, st_counts, flag_counts
