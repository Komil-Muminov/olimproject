"""Выводы по своду реестров: статус ИНН по периодам, кто пропал/появился, резкий рост/падение, поиск."""
from dataclasses import dataclass

from registry_stats import INN_RE

JUMP = 2.0      # «резко» — сумма выросла хотя бы вдвое или упала хотя бы вдвое к прошлому периоду


@dataclass
class PartyRow:
    inn: str
    name: str
    sums: list          # по периодам (в порядке времени)
    counts: list
    total: float
    status: str


def party_row(p, reg) -> PartyRow:
    order = reg.order
    sums = [round(p.sums.get(f, 0.0), 2) for f in order]
    counts = [p.counts.get(f, 0) for f in order]
    return PartyRow(p.inn, p.name, sums, counts, round(sum(sums), 2), status(counts, reg))


def status(counts: list, reg) -> str:
    """Где ИНН встречается: весь период / пропал после … / появился с … / с перерывом."""
    if len(counts) < 2:
        return ""
    here = [c > 0 for c in counts]
    labels = [reg.files[f].label for f in reg.order]
    if all(here):
        return "весь период"
    if here[0] and not here[-1]:
        last = max(i for i, h in enumerate(here) if h)
        return f"пропал после «{labels[last]}»"
    if not here[0] and here[-1]:
        first = min(i for i, h in enumerate(here) if h)
        return f"появился с «{labels[first]}»"
    return "с перерывом"


def party_rows(book: dict, reg) -> list:
    """Все плательщики (или получатели), крупные сверху."""
    return sorted((party_row(p, reg) for p in book.values()), key=lambda r: -r.total)


def gone_and_new(rows: list) -> tuple:
    """(пропали, появились) — только с правильным ИНН («0» и т.п. — это много разных организаций)."""
    ok = [r for r in rows if INN_RE.match(r.inn)]
    return ([r for r in ok if r.status.startswith("пропал")],
            [r for r in ok if r.status.startswith("появился")])


def jumps(rows: list, reg) -> list:
    """Резкий рост или падение между соседними периодами: (строка, «с», «на», было, стало, разница).
    Крупные изменения сверху."""
    labels = [reg.files[f].label for f in reg.order]
    out = []
    for r in rows:
        if not INN_RE.match(r.inn):
            continue
        for i in range(1, len(r.sums)):
            a, b = r.sums[i - 1], r.sums[i]
            if r.counts[i - 1] and r.counts[i] and a > 0 and b > 0 and (b >= a * JUMP or b * JUMP <= a):
                out.append((r, labels[i - 1], labels[i], a, b, round(b - a, 2)))
    return sorted(out, key=lambda x: -abs(x[5]))


def search(reg, inn: str) -> dict:
    """ИНН как плательщик и как получатель: {"payer": PartyRow|None, "rec": PartyRow|None}."""
    inn = "".join(inn.split())
    if inn.isdigit() and len(inn) < 9:      # потерялся ноль в начале
        inn = inn.zfill(9)
    return {side: (party_row(book[inn], reg) if inn in book else None)
            for side, book in (("payer", reg.payers), ("rec", reg.recs))}
