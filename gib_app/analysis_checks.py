"""Проверки по всей таблице организаций: одно название — разные ИНН; один ИНН — разные названия."""
from collections import defaultdict

from analysis_model import BAD_INN, SAME_NAME
from inn_check import _name_key

MIN_NAME = 4        # короче — слишком общее название, совпадения ничего не значат
SHOW_NAMES = 20     # в ячейке — первые названия, остальные числом (в ячейку Excel влезает 32 767 знаков)


def mark_same_names(entities: list) -> list:
    """Одинаковое название (без кавычек, пробелов и знаков) у разных ИНН одной роли: помечает организации
    флагом SAME_NAME. → строки группы «Одно название — разные ИНН»: [№ группы, название, кто, ИНН, всего,
    записей]; крупные группы сверху. Организации с неправильным ИНН («0» и т.п.) не учитываются."""
    groups = defaultdict(list)
    for e in entities:
        key = _name_key(e.name)
        if len(key) >= MIN_NAME and BAD_INN not in e.flags:
            groups[(e.role, key)].append(e)
    found = [g for g in groups.values() if len({e.inn.lstrip("0") for e in g}) > 1]
    found.sort(key=lambda g: -max(e.total for e in g))
    rows = []
    for n, group in enumerate(found, 1):
        for e in sorted(group, key=lambda e: -e.total):
            e.flags.add(SAME_NAME)
            rows.append([n, e.name, e.role, e.inn, e.total, sum(e.counts)])
    return rows


def distinct_names(names) -> list:
    """Разные написания названия (без учёта кавычек, регистра и пробелов) — самое частое первым."""
    seen, out = set(), []
    for n in (sorted(names, key=lambda x: -names[x]) if hasattr(names, "most_common") else names):
        if (key := _name_key(n)) and key not in seen:
            seen.add(key)
            out.append(" ".join(str(n).split()))
    return out


def inn_names(entities: list) -> list:
    """Один ИНН — разные названия: [ИНН, кто, сколько названий, названия через « | », всего, записей];
    больше всего вариантов — сверху. Неправильные ИНН («0») не в счёт."""
    found = [e for e in entities if len(e.variants) > 1 and BAD_INN not in e.flags]
    found.sort(key=lambda e: (-len(e.variants), -e.total))
    more = lambda e: f" | …и ещё {len(e.variants) - SHOW_NAMES}" if len(e.variants) > SHOW_NAMES else ""
    return [[e.inn, e.role, len(e.variants), " | ".join(e.variants[:SHOW_NAMES]) + more(e), e.total, sum(e.counts)]
            for e in found]
