"""Поиск ИНН в своде реестров: быстрый фильтр по ИНН / части названия и разбор введённого текста."""
import re
from bisect import bisect_right

from registry_analysis import party_rows


def build_index(reg) -> dict:
    """Для фильтра по каждой стороне: (строки PartyRow — крупные сверху, все «инн название» строчными
    одной строкой через перевод строки, начала строк). Поиск по одной большой строке в разы быстрее."""
    out = {}
    for side, book in (("payer", reg.payers), ("rec", reg.recs)):
        rows = party_rows(book, reg)
        keys = [f"{r.inn} {r.name.lower()}" for r in rows]
        starts, pos = [], 0
        for k in keys:
            starts.append(pos)
            pos += len(k) + 1
        out[side] = (rows, "\n".join(keys), starts)
    return out


def matching(index_side, q: str) -> list:
    """Строки, где «инн название» содержит q (в порядке убывания суммы)."""
    rows, text, starts = index_side
    if not q:
        return rows
    found, last = [], -1
    for m in re.finditer(re.escape(q), text):
        i = bisect_right(starts, m.start()) - 1
        if i != last:
            found.append(rows[i])
            last = i
    return found


def clean_inn_query(raw: str, reg=None) -> str:
    """Чистый ИНН из введённого текста; по названию — самая крупная по сумме подходящая организация."""
    raw = raw.strip()
    m = re.match(r"^(\d{9,12})", raw)
    if m:
        return m.group(1)
    digits = "".join(c for c in raw if c.isdigit())
    if len(digits) in (9, 10, 12):
        return digits
    if reg and not digits:
        q = raw.lower()
        hits = [p for book in (reg.payers, reg.recs) for p in book.values() if q in p.name.lower()]
        if hits:
            return max(hits, key=lambda p: p.total).inn
    return digits or raw
