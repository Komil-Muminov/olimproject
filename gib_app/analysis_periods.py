"""Как разбить реестры на части: по файлам (как есть), по кварталам или по месяцам (по дате документа).

Слот — что входит в часть: ("f", файл) или ("m", (год, месяц)). Платежи с датой вне периодов файлов
(«странная дата») в кварталы и месяцы не попадают — они видны на листе ошибок."""
import os

from registry_stats import MONTHS

FILES, QUARTERS, MONTHS_MODE = "files", "quarters", "months"
MODE_TITLES = {FILES: "по файлам", QUARTERS: "по кварталам", MONTHS_MODE: "по месяцам"}


def all_months(reg) -> list:
    """Основные месяцы всех файлов (без «чужих» дат), по порядку."""
    return sorted({m for info in reg.files for m in info.main_months})


def available_modes(reg) -> list:
    """Разбивка по кварталам / месяцам имеет смысл, если в реестрах больше одного квартала / месяца."""
    months = all_months(reg)
    modes = [FILES]
    if len({(y, (m - 1) // 3) for y, m in months}) > 1:
        modes.append(QUARTERS)
    if len(months) > 1:
        modes.append(MONTHS_MODE)
    return modes


def registry_slots(reg, mode: str) -> list:
    """→ [(подпись, файлы, порядок (год, месяц, №), [слоты])] — части реестров в порядке времени."""
    if mode == FILES:
        out = []
        for f in reg.order:
            info = reg.files[f]
            many = sum(i.path == info.path for i in reg.files) > 1
            y, m = info.main_months[0] if info.main_months else (9999, 0)
            out.append((info.label + (f" · {info.sheet}" if many else ""), os.path.basename(info.path),
                        (y, m, f), [("f", f)]))
        return out
    groups = {}
    for y, m in all_months(reg):
        key = (y, (m - 1) // 3) if mode == QUARTERS else (y, m)
        groups.setdefault(key, []).append((y, m))
    out = []
    for key, months in sorted(groups.items()):
        y, m = months[0]
        label = f"{key[1] + 1} кв. {y}" if mode == QUARTERS else f"{MONTHS[m - 1]} {y}"
        files = sorted({os.path.basename(i.path) for i in reg.files if set(months) & set(i.main_months)})
        out.append((label, ", ".join(files), (y, m, 0), [("m", ym) for ym in months]))
    return out


def party_totals(party, slots: list) -> tuple:
    """Суммы и количество платежей организации по частям."""
    sums, counts = [], []
    for *_, keys in slots:
        s = c = 0
        for kind, k in keys:
            if kind == "f":
                s, c = s + party.sums.get(k, 0.0), c + party.counts.get(k, 0)
            else:
                s, c = s + party.msums.get(k, 0.0), c + party.mcounts.get(k, 0)
        sums.append(round(s, 2))
        counts.append(c)
    return sums, counts


def pair_totals(reg, slots: list) -> dict:
    """Кто кому платит: (ИНН плательщика, ИНН получателя) → (суммы по частям, платежей по частям)."""
    index = {(kind, k): i for i, (*_, keys) in enumerate(slots) for kind, k in keys}
    out = {}
    for book, kind in ((reg.pairs_f, "f"), (reg.pairs_m, "m")):
        for (a, b, k), (s, c) in book.items():
            i = index.get((kind, k))
            if i is None:
                continue
            sums, counts = out.setdefault((a, b), ([0.0] * len(slots), [0] * len(slots)))
            sums[i] += s
            counts[i] += c
    return out
