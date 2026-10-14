"""Реестры платежей → части и организации глубокого анализа (плательщики и получатели)."""
import os
import re

from analysis_model import BAD_INN, DUP_PAY, ODD_DATE, PAYER, RECIPIENT, Entity, Part
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


def _pair(a, b) -> str:
    return " → ".join(str(x) for x in (a, b) if x not in (None, ""))


def registry_errors(reg) -> tuple:
    """→ (строки «Ошибки подробно», {флаг: {(роль, ИНН)}} — у кого какая ошибка)."""
    fname = lambda f: os.path.basename(reg.files[f].path)
    rows, marked = [], {DUP_PAY: set(), ODD_DATE: set()}
    for f, p, first in reg.dups:
        marked[DUP_PAY].update(_sides(p))
        rows.append([DUP_PAY, fname(f), f"строка {p.row}", _pair(p.payer_inn, p.rec_inn), _pair(p.payer, p.rec),
                     p.doc, p.date, p.amount, f"такой же платёж: {fname(first[0])}, строка {first[1]}"])
    for f, p, side, inn in reg.bad_inn:
        name = p.payer if side == PAYER else p.rec
        hint = inn_in_name(name)
        rows.append([BAD_INN, fname(f), f"строка {p.row}", inn or "(пусто)", name, p.doc, p.date, p.amount,
                     f"ИНН {side}я" + (f"; в названии указан ИНН {hint}" if hint else "")])
    for f, p in reg.odd_dates:
        marked[ODD_DATE].update(_sides(p))
        rows.append([ODD_DATE, fname(f), f"строка {p.row}", _pair(p.payer_inn, p.rec_inn), _pair(p.payer, p.rec),
                     p.doc, p.date, p.amount, f"период файла: {reg.files[f].label}"])
    return rows, marked


def registry_entities(reg, marked: dict) -> list:
    out = []
    for book, role in ((reg.payers, PAYER), (reg.recs, RECIPIENT)):
        for inn, party in book.items():
            e = Entity(inn, party.name, role, [round(party.sums.get(f, 0.0), 2) for f in reg.order],
                       [party.counts.get(f, 0) for f in reg.order], names=max(1, len({_name_key(n) for n in party.names})))
            if not INN_RE.match(inn):
                e.flags.add(BAD_INN)
            e.flags.update(flag for flag in (DUP_PAY, ODD_DATE) if (role, inn) in marked[flag])
            out.append(e)
    return out
