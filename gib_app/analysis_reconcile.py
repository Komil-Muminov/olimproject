"""Сверка списка с реестрами: каждая организация из обычных таблиц (например, «30 фоиза») — в одной строке
отмеченные столбцы списка и её платежи по реестрам (без выводов — смотрит пользователь). Связь — по ИНН
(без нулей впереди), а если в списке ИНН нет — по названию."""
from dataclasses import dataclass, field

from analysis_list_map import org_key
from analysis_model import BY_NAME, ORG, PAYER, RECIPIENT
from compare_finder import match_key


@dataclass
class ReconRow:
    inn: str
    name: str
    where: str          # в каких частях списка есть
    values: list        # отмеченные столбцы списка — по порядку list_heads (None — нет в этой части)
    paid: list          # заплатил — по каждой части реестров
    payments: int       # сколько платежей
    got: float          # сколько получил (если был получателем)
    to: list = field(default_factory=list)      # кому платил: [(название, ИНН, [суммы по частям], [платежей])]
    frm: list = field(default_factory=list)     # от кого получал: [(название, ИНН, [суммы], [платежей])]


def _links(data, reg: list, keys: set) -> tuple:
    """Связки «кто кому платит» только для организаций из списка: (кому платил, от кого получал)."""
    to, frm = {}, {}
    for x in data.pairs:
        entry = lambda name, inn: (name, inn, [x.sums[i] for i in reg], [x.counts[i] for i in reg])
        if match_key(x.payer_inn) in keys:
            to.setdefault(match_key(x.payer_inn), []).append(entry(x.rec, x.rec_inn))
        if match_key(x.rec_inn) in keys:
            frm.setdefault(match_key(x.rec_inn), []).append(entry(x.payer, x.payer_inn))
    return to, frm


def reconcile(data) -> tuple:
    """→ (части реестров, [(часть списка, заголовок столбца)], [ReconRow]) — пусто, если нет списков или
    нет реестров. Строки — в порядке списка."""
    reg = [i for i, p in enumerate(data.parts) if p.kind == "payments"]
    lists = [i for i, p in enumerate(data.parts) if p.kind == "list"]
    if not reg or not lists:
        return [], [], []
    books, by_name = {PAYER: {}, RECIPIENT: {}}, {PAYER: {}, RECIPIENT: {}}
    for e in data.entities:
        if e.role in books:
            books[e.role][match_key(e.inn)] = e
            by_name[e.role].setdefault(org_key(e.name), e)
    orgs = [e for e in data.entities if e.role == ORG]
    found = {}                          # организация списка → (плательщик, получатель) в реестрах
    for e in orgs:                      # без ИНН — по названию (без кавычек и форм)
        if BY_NAME in e.flags:
            found[id(e)] = tuple(by_name[role].get(org_key(e.name)) for role in (PAYER, RECIPIENT))
        else:
            found[id(e)] = tuple(books[role].get(match_key(e.inn)) for role in (PAYER, RECIPIENT))
    to, frm = _links(data, reg, {match_key(x.inn) for pair in found.values() for x in pair if x})
    heads = [(data.parts[i], h) for i in lists for h in data.parts[i].heads]
    rows = []
    for e in orgs:
        p, r = found[id(e)]
        values = [v for i in lists for v in (e.cols[i] or [None] * len(data.parts[i].heads))]
        rows.append(ReconRow(
            e.inn, e.name, ", ".join(data.parts[i].label for i in lists if e.counts[i]), values,
            [round(p.sums[i], 2) if p and p.counts[i] else 0.0 for i in reg],
            sum(p.counts[i] for i in reg) if p else 0,
            round(sum(r.sums[i] for i in reg), 2) if r else 0.0,
            to.get(match_key(p.inn), []) if p else [], frm.get(match_key(r.inn), []) if r else []))
    return [data.parts[i] for i in reg], heads, rows
