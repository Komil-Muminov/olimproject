"""Сверка списка с реестрами: каждая организация из обычных таблиц (например, «30 фоиза») — платила ли она
по реестрам платежей, сколько и в каком периоде. Связь — по ИНН (без нулей впереди)."""
from dataclasses import dataclass

from analysis_model import ORG, PAYER, RECIPIENT
from compare_finder import match_key

PAID, NOT_PAID = "Платил", "Не платил"


@dataclass
class ReconRow:
    inn: str
    name: str
    where: str          # в каких частях списка есть
    paid: list          # заплатил — по каждой части реестров
    payments: int       # сколько платежей
    got: float          # сколько получил (если был получателем)

    @property
    def result(self) -> str:
        return PAID if self.payments else NOT_PAID


def reconcile(data) -> tuple:
    """→ (части реестров, [ReconRow]) — пусто, если нет списков или нет реестров. «Не платил» — сверху."""
    reg = [i for i, p in enumerate(data.parts) if p.kind == "payments"]
    lists = [i for i, p in enumerate(data.parts) if p.kind == "list"]
    if not reg or not lists:
        return [], []
    books = {PAYER: {}, RECIPIENT: {}}
    for e in data.entities:
        if e.role in books:
            books[e.role][match_key(e.inn)] = e
    rows = []
    for e in data.entities:
        if e.role != ORG:
            continue
        p, r = books[PAYER].get(match_key(e.inn)), books[RECIPIENT].get(match_key(e.inn))
        rows.append(ReconRow(
            e.inn, e.name, ", ".join(data.parts[i].label for i in lists if e.counts[i]),
            [round(p.sums[i], 2) if p and p.counts[i] else 0.0 for i in reg],
            sum(p.counts[i] for i in reg) if p else 0,
            round(sum(r.sums[i] for i in reg), 2) if r else 0.0))
    rows.sort(key=lambda x: (x.result == PAID, -sum(x.paid), x.name))
    return [data.parts[i] for i in reg], rows
