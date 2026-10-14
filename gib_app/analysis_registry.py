"""Реестры платежей → части и организации единой таблицы анализа (плательщики и получатели)."""
import os
import re

from analysis_model import BAD_INN, DUP_PAY, NAMES, ODD_DATE, PAYER, RECIPIENT, Entity, Part
from inn_check import _name_key
from registry_stats import INN_RE

INN_IN_TEXT = re.compile(r"(?<!\d)\d{9}(?!\d)")


def inn_in_name(name: str) -> str:
    """ИНН, вписанный в название («… ИНН 010010058»), — подсказка, каким он должен быть."""
    m = INN_IN_TEXT.search(name or "")
    return m.group(0) if m else ""


def registry_parts(reg) -> list:
    parts = []
    for f in reg.order:
        info = reg.files[f]
        many = sum(i.path == info.path for i in reg.files) > 1
        y, m = info.main_months[0] if info.main_months else (9999, 0)
        parts.append(Part(info.label + (f" · {info.sheet}" if many else ""), os.path.basename(info.path),
                          "payments", (y, m, f)))
    return parts


def _sides(p):
    """(роль, ИНН) сторон платежа, которые есть в документе."""
    return [(role, inn) for role, inn in ((PAYER, p.payer_inn), (RECIPIENT, p.rec_inn)) if inn is not None]


def registry_entities(reg) -> tuple:
    """→ (организации, подробности ошибок {(флаг, "payments"): [(ключ, строка)]})."""
    fname = lambda f: os.path.basename(reg.files[f].path)
    details = {DUP_PAY: [], BAD_INN: [], ODD_DATE: []}
    marked = {DUP_PAY: set(), ODD_DATE: set()}
    for f, p, first in reg.dups:
        row = [fname(f), p.row, p.doc, p.date, p.payer, p.payer_inn, p.rec, p.rec_inn, p.amount,
               fname(first[0]), first[1]]
        for key in _sides(p):
            marked[DUP_PAY].add(key)
            details[DUP_PAY].append((key, row))
    for f, p, side, inn in reg.bad_inn:
        name = p.payer if side == PAYER else p.rec
        details[BAD_INN].append(((side, inn), [fname(f), p.row, side, inn or "(пусто)", name, inn_in_name(name),
                                               p.doc, p.date, p.amount]))
    for f, p in reg.odd_dates:
        row = [fname(f), reg.files[f].label, p.row, p.doc, p.date, p.payer, p.rec, p.amount]
        for key in _sides(p):
            marked[ODD_DATE].add(key)
            details[ODD_DATE].append((key, row))
    out = []
    for book, role in ((reg.payers, PAYER), (reg.recs, RECIPIENT)):
        for inn, party in book.items():
            e = Entity(inn, party.name, role, [round(party.sums.get(f, 0.0), 2) for f in reg.order],
                       [party.counts.get(f, 0) for f in reg.order])
            if not INN_RE.match(inn):
                e.flags.add(BAD_INN)
            if len({_name_key(n) for n in party.names}) > 1:
                e.flags.add(NAMES)
            for flag in (DUP_PAY, ODD_DATE):
                if (role, inn) in marked[flag]:
                    e.flags.add(flag)
            out.append(e)
    return out, {(flag, "payments"): rows for flag, rows in details.items()}
