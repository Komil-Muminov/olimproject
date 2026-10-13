"""Сравнение двух списков по ИНН / РМА («Муқоиса»)."""
from openpyxl.utils import get_column_letter

from compare_finder import compare_keyword, find_lists
from compare_models import Diff, metric_key, metric_title
from utils import cell_val, norm_text, to_num_safe

# Обратная совместимость для приватных импортов
_cell = cell_val
_metric = metric_key
_title = metric_title
num = to_num_safe


def compare(data: list[list]):
    """→ (Diff, сторона1, сторона2) или None, если на листе нет двух списков с ИНН."""
    found = find_lists(data)
    if not found:
        return None
    header, (a, b) = found
    pairs = []
    used_b = set()
    for ca in a.value_cols:
        ma = metric_key(cell_val(data, header, ca))
        if not ma:
            continue
        for cb in b.value_cols:
            if cb not in used_b and metric_key(cell_val(data, header, cb)) == ma:
                pairs.append((metric_title(cell_val(data, header, ca)), ca, cb))
                used_b.add(cb)
                break

    def name(side, k):
        return str(cell_val(data, side.rows[k], side.name_col) or "").strip() if side.name_col else ""

    def values(side, k, cols):
        return {str(cell_val(data, header, c) or get_column_letter(c)).replace("\n", " ").strip():
                cell_val(data, side.rows[k], c) for c in cols}

    only1 = [(a.keys[k], name(a, k), values(a, k, [p[1] for p in pairs] or a.value_cols[:2]))
             for k in a.rows if k not in b.rows]
    only2 = [(b.keys[k], name(b, k), values(b, k, [p[2] for p in pairs] or b.value_cols[:2]))
             for k in b.rows if k not in a.rows]
    changed, same, names_differ = [], 0, 0
    org_changed, org_same, org_names = [], [], []
    common = [k for k in b.rows if k in a.rows]
    for k in common:
        if norm_text(name(a, k)).replace('"', "") != norm_text(name(b, k)).replace('"', ""):
            names_differ += 1
            org_names.append((b.keys[k], name(a, k), name(b, k)))
        diff_here, vals = False, []
        for title, ca, cb in pairs:
            v1, v2 = to_num_safe(cell_val(data, a.rows[k], ca)), to_num_safe(cell_val(data, b.rows[k], cb))
            vals.append((v1, v2))
            if abs(v2 - v1) > 0.005:
                changed.append((b.keys[k], name(b, k) or name(a, k), title, v1, v2, round(v2 - v1, 2)))
                diff_here = True
        (org_changed if diff_here else org_same).append((b.keys[k], name(b, k) or name(a, k), vals))
        same += not diff_here
    org_only1 = [(a.keys[k], name(a, k), [to_num_safe(cell_val(data, a.rows[k], p[1])) for p in pairs])
                 for k in a.rows if k not in b.rows]
    org_only2 = [(b.keys[k], name(b, k), [to_num_safe(cell_val(data, b.rows[k], p[2])) for p in pairs])
                 for k in b.rows if k not in a.rows]
    duplicates = [(i, side.keys[k], name(side, k), rows)
                  for i, side in ((1, a), (2, b)) for k, rows in side.all_rows.items() if len(rows) > 1]
    totals = [(title, round(sum(to_num_safe(cell_val(data, r, ca)) for r in a.rows.values()), 2),
               round(sum(to_num_safe(cell_val(data, r, cb)) for r in b.rows.values()), 2)) for title, ca, cb in pairs]
    return Diff(only1, only2, changed, same, len(common), names_differ, pairs, duplicates, totals,
                org_only1, org_only2, org_changed, org_same, org_names), a, b


def find_compares(sheets: dict):
    """Все листы сравнения в книге (в порядке листов): название с ключевым словом и два списка с ИНН."""
    found = []
    for name, data in sheets.items():
        if compare_keyword(name):
            res = compare(data)
            if res:
                found.append((name, *res))
    return found


def best_compare(found: list):
    """Какой лист предложить первым: тот, где можно сравнить суммы."""
    return max(range(len(found)), key=lambda i: (len(found[i][1].pairs), i)) if found else None


def max_change(vals: list) -> float:
    """Наибольшая по модулю разница среди показателей: [(было, стало)] → число."""
    return max((abs((v2 or 0) - (v1 or 0)) for v1, v2 in vals), default=0.0)


def sorted_changes(diff: Diff):
    """Организации с изменившейся суммой — крупные разницы сверху."""
    return sorted(diff.org_changed, key=lambda o: -max_change(o[2]))
