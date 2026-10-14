"""Реестры платежей → части, организации, пары «кто кому платит» и ошибки глубокого анализа."""
import os
import re

from analysis_model import BAD_INN, DUP_PAY, ODD_DATE, PAYER, RECIPIENT, Entity, Pair, Part
from analysis_periods import pair_totals, party_totals
from inn_check import _name_key
from registry_stats import INN_DIGITS, inn_ok

INN_IN_TEXT = re.compile(r"(?<!\d)\d{9}(?!\d)")
SIDE_OF = {PAYER: "плательщика", RECIPIENT: "получателя"}


def inn_in_name(name: str) -> str:
    """ИНН, вписанный в название («… ИНН 010010058»), — подсказка, каким он должен быть."""
    m = INN_IN_TEXT.search(name or "")
    return m.group(0) if m else ""


def inn_problem(inn: str) -> str:
    """Что не так с ИНН — простыми словами."""
    digits = re.sub(r"\D", "", inn or "")
    if not inn:
        return "ИНН пустой"
    if digits and not digits.strip("0"):
        return "вместо ИНН стоит 0 (не указан)"
    signs = " (есть буквы или знаки)" if digits != inn else ""
    tail = f" — возможно, потерян 0 в начале (должно быть {digits.zfill(INN_DIGITS)})" if len(digits) == 8 else ""
    return f"{len(digits)} цифр вместо {INN_DIGITS}{signs}{tail}"


def registry_parts(slots: list) -> list:
    return [Part(label, source, "payments", order) for label, source, order, _ in slots]


def _sides(p):
    """(роль, ИНН) сторон платежа, которые есть в документе."""
    return [(role, inn) for role, inn in ((PAYER, p.payer_inn), (RECIPIENT, p.rec_inn)) if inn is not None]


def registry_errors(reg) -> tuple:
    """→ ({группа ошибок: [строки]} для листов ошибок, {флаг: {(роль, ИНН)}} — у кого какая)."""
    fname = lambda f: os.path.basename(reg.files[f].path)
    groups, marked = {"dup": [], "bad_pay": [], "odd": []}, {DUP_PAY: set(), ODD_DATE: set()}
    for n, (f, p, (f0, row0)) in enumerate(reg.dups, 1):
        marked[DUP_PAY].update(_sides(p))
        # пара строк: где платёж записан впервые и где повтор (данные те же — отличается только строка)
        tail = [p.payer, p.payer_inn, p.rec, p.rec_inn, p.purpose]
        groups["dup"] += [[n, fname(f0), row0, p.doc, p.date, p.amount, "платёж"] + tail,
                          [n, fname(f), p.row, p.doc, p.date, p.amount, "повтор"] + tail]
    for f, p, side, inn in reg.bad_inn:
        name = p.payer if side == PAYER else p.rec
        groups["bad_pay"].append([fname(f), p.row, p.doc, p.date, p.amount, SIDE_OF[side], name,
                                  inn or "(пусто)", inn_problem(inn), inn_in_name(name)])
    for f, p in reg.odd_dates:
        marked[ODD_DATE].update(_sides(p))
        groups["odd"].append([fname(f), p.row, p.doc, p.date, p.amount, reg.files[f].label, p.payer, p.payer_inn,
                              p.rec, p.rec_inn])
    return groups, marked


def registry_entities(reg, marked: dict, slots: list) -> list:
    out = []
    for book, role in ((reg.payers, PAYER), (reg.recs, RECIPIENT)):
        for inn, party in book.items():
            sums, counts = party_totals(party, slots)
            e = Entity(inn, party.name, role, sums, counts, names=max(1, len({_name_key(n) for n in party.names})))
            if not inn_ok(inn):
                e.flags.add(BAD_INN)
            e.flags.update(flag for flag in (DUP_PAY, ODD_DATE) if (role, inn) in marked[flag])
            out.append(e)
    return out


def registry_pairs(reg, slots: list) -> list:
    """Кто кому платит — все пары плательщик → получатель."""
    name = lambda book, inn: book[inn].name if inn in book else ""
    return [Pair(a, name(reg.payers, a), b, name(reg.recs, b), [round(x, 2) for x in sums], counts)
            for (a, b), (sums, counts) in pair_totals(reg, slots).items()]
