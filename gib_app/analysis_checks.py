"""Проверки по всей таблице организаций: одно название — разные ИНН."""
from collections import defaultdict

from analysis_model import BAD_INN, SAME_NAME
from inn_check import _name_key

MIN_NAME = 4        # короче — слишком общее название, совпадения ничего не значат


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
