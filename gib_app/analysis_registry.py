"""Реестры платежей → части, организации и ошибки глубокого анализа (плательщики и получатели)."""
import os
import re

from analysis_model import BAD_INN, DUP_PAY, ODD_DATE, PAYER, RECIPIENT, Entity, Part
from inn_check import _name_key
from registry_stats import INN_RE

INN_IN_TEXT = re.compile(r"(?<!\d)\d{9}(?!\d)")
SIDE_OF = {PAYER: "плательщика", RECIPIENT: "получателя"}


def inn_in_name(name: str) -> str:
    """ИНН, вписанный в название («… ИНН 010010058»), — подсказка, каким он должен быть."""
    m = INN_IN_TEXT.search(name or "")
    return m.group(0) if m else ""


def inn_problem(inn: str) -> str:
    """Что не так с ИНН — простыми словами."""
    if not inn:
        return "ИНН пустой"
    if inn.isdigit() and not inn.strip("0"):
        return "вместо ИНН стоит 0 (не указан)"
    if not inn.isdigit():
        return "в ИНН есть буквы или знаки"
    tail = " — возможно, потерян 0 в начале (должно быть " + inn.zfill(9) + ")" if len(inn) == 8 else ""
    return f"{len(inn)} цифр вместо 9{tail}"


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


def registry_errors(reg) -> tuple:
    """→ ({группа ошибок: [строки]} для листа «Ошибки в реестрах», {флаг: {(роль, ИНН)}} — у кого какая)."""
    fname = lambda f: os.path.basename(reg.files[f].path)
    groups, marked = {"dup": [], "bad_pay": [], "odd": []}, {DUP_PAY: set(), ODD_DATE: set()}
    for f, p, (f0, row0) in reg.dups:
        marked[DUP_PAY].update(_sides(p))
        where = f"строка {row0}" if f0 == f else f"{fname(f0)}, строка {row0}"
        groups["dup"].append([fname(f), p.row, p.doc, p.date, p.amount, where, p.payer, p.payer_inn, p.rec,
                              p.rec_inn, p.purpose])
    for f, p, side, inn in reg.bad_inn:
        name = p.payer if side == PAYER else p.rec
        groups["bad_pay"].append([fname(f), p.row, p.doc, p.date, p.amount, SIDE_OF[side], name,
                                  inn or "(пусто)", inn_problem(inn), inn_in_name(name)])
    for f, p in reg.odd_dates:
        marked[ODD_DATE].update(_sides(p))
        groups["odd"].append([fname(f), p.row, p.doc, p.date, p.amount, reg.files[f].label, p.payer, p.payer_inn,
                              p.rec, p.rec_inn])
    return groups, marked


def registry_entities(reg, marked: dict) -> list:
    out = []
    for book, role in ((reg.payers, PAYER), (reg.recs, RECIPIENT)):
        for inn, party in book.items():
            e = Entity(inn, party.name, role, [round(party.sums.get(f, 0.0), 2) for f in reg.order],
                       [party.counts.get(f, 0) for f in reg.order],
                       names=max(1, len({_name_key(n) for n in party.names})))
            if not INN_RE.match(inn):
                e.flags.add(BAD_INN)
            e.flags.update(flag for flag in (DUP_PAY, ODD_DATE) if (role, inn) in marked[flag])
            out.append(e)
    return out
