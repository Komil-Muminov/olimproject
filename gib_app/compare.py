"""Сравнение двух списков по ИНН / РМА («Муқоиса»).

Лист, в названии которого есть ключевое слово сравнения (муқоиса, фарқият, сравнение,
compare…), содержит два списка организаций рядом. Программа находит в каждом столбец
ИНН/РМА, название и суммы и показывает различия: новые, пропавшие, изменившиеся суммы.
Файл не меняется.
"""
import re
from collections import Counter
from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

# Корни слов «сравнение / различие» — таджикский (без особых букв: қ→к, ҳ→х, ӣ→и, ҷ→ч, ғ→г, ӯ→у),
# русский, английский. Ищутся в названии листа в любом месте, любой регистр.
COMPARE_STEMS = (
    "мукоис", "киёс", "фарк", "тафовут", "мутобик", "тагйир", "санчиш", "баробаркун", "хамоханг",
    "сравн", "различ", "отлич", "разниц", "сверк", "сверить", "сопостав", "расхожд", "изменени", "дельта",
    "compar", "diff", "delta", "reconcil", "change", "match", "versus",
)
TJ_TO_RU = str.maketrans("қҳӣҷғӯҚҲӢҶҒӮ", "кхичгуКХИЧГУ")
KEY_RE = re.compile(r"^\d{8,14}$")


def norm_text(v) -> str:
    return " ".join(str(v or "").translate(TJ_TO_RU).lower().split())


def compare_keyword(sheet_name: str):
    """Ключевое слово сравнения в названии листа (или None)."""
    t = norm_text(sheet_name)
    for stem in COMPARE_STEMS:
        if stem in t:
            return stem
    if re.search(r"(^|[^a-z])vs([^a-z]|$)", t):
        return "vs"
    return None


def key_of(v):
    """ИНН/РМА: 8–14 цифр (ведущие нули могли потеряться, если Excel хранит число)."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s if KEY_RE.match(s) else None


def match_key(k):
    return k.lstrip("0")


def num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


@dataclass
class SideList:
    key_col: int
    name_col: int | None
    cols: list                         # столбцы списка (для подписи)
    value_cols: list                   # числовые столбцы
    rows: dict = field(default_factory=dict)   # ключ сравнения -> строка
    keys: dict = field(default_factory=dict)   # ключ сравнения -> ИНН как в файле
    all_rows: dict = field(default_factory=dict)   # ключ сравнения -> все строки (для дубликатов)
    year: str | None = None

    def label(self, i):
        span = f"{get_column_letter(min(self.cols))}–{get_column_letter(max(self.cols))}"
        year = f" · {self.year}" if self.year else ""
        return f"Список {i}{year} — {len(self.rows)} орг. (столбцы {span})"


@dataclass
class Diff:
    only1: list          # (ИНН, название, {показатель: значение})
    only2: list
    changed: list        # (ИНН, название, показатель, было, стало, разница)
    same: int            # организаций без изменений
    common: int
    names_differ: int
    pairs: list          # [(показатель, столбец1, столбец2)]
    duplicates: list = field(default_factory=list)   # (номер списка, ИНН, название, [строки])
    totals: list = field(default_factory=list)       # (показатель, итог списка 1, итог списка 2)
    # по организациям, значения в порядке pairs: (ИНН, название, [значения]) / (ИНН, название, [(было, стало)])
    org_only1: list = field(default_factory=list)
    org_only2: list = field(default_factory=list)
    org_changed: list = field(default_factory=list)
    org_same: list = field(default_factory=list)
    org_names: list = field(default_factory=list)    # (ИНН, название в списке 1, в списке 2) — записаны по-разному


def find_lists(data):
    """Два списка на листе: столбцы, где почти все значения — ИНН/РМА."""
    width = max((len(r) for r in data), default=0)
    key_cols = []
    for c in range(1, width + 1):
        vals = [r[c - 1] for r in data if c - 1 < len(r) and r[c - 1] not in (None, "")]
        keys = [v for v in vals if key_of(v)]
        if len(keys) >= 10 and len(keys) >= 0.6 * len(vals):
            key_cols.append(c)
    if len(key_cols) < 2:
        return None
    key_cols = key_cols[:2]

    first = min(r for r, row in enumerate(data, 1) for c in key_cols
                if c - 1 < len(row) and key_of(row[c - 1]))
    # шапка — ближайшая строка выше данных с текстом (итоговую строку «Ҳамагӣ» пропускаем)
    header = max(1, first - 1)
    for r in range(first - 1, 0, -1):
        row = data[r - 1]
        if any(isinstance(v, str) and len(v.strip()) > 1 and not norm_text(v).startswith(("хамаги", "итого", "всего"))
               for v in row):
            header = r
            break

    # блоки столбцов, разделённые пустыми столбцами
    def used(c):
        return any(c - 1 < len(row) and row[c - 1] not in (None, "") for row in data[header - 1:])

    blocks, cur = [], []
    for c in range(1, width + 1):
        if used(c):
            cur.append(c)
        elif cur:
            blocks.append(cur)
            cur = []
    if cur:
        blocks.append(cur)

    def last_row(cols):
        return max((r for r, row in enumerate(data, 1) for c in cols if c - 1 < len(row) and row[c - 1] not in (None, "")),
                   default=0)

    sides = []
    for k in key_cols:
        block = next(b for b in blocks if k in b)
        sides.append(SideList(k, None, list(block), []))
    # блоки без ИНН (суммы отдельно от списка) — к списку с той же длиной
    for b in blocks:
        if any(k in b for k in key_cols):
            continue
        lr = last_row(b)
        target = min(sides, key=lambda s: abs(last_row([s.key_col]) - lr))
        target.cols.extend(b)

    for s in sides:
        s.cols.sort()
        texts = Counter()
        for c in s.cols:
            if c == s.key_col:
                continue
            col_vals = [row[c - 1] for row in data[header:] if c - 1 < len(row) and row[c - 1] not in (None, "")]
            if not col_vals:
                continue
            nums = [v for v in col_vals if num(v) is not None]
            strs = [v for v in col_vals if isinstance(v, str) and len(v.strip()) > 3 and not key_of(v)]
            if len(strs) > 0.5 * len(col_vals):
                texts[c] = len(strs)
            elif len(nums) > 0.5 * len(col_vals) and not _is_numbering(col_vals):
                s.value_cols.append(c)
        s.name_col = texts.most_common(1)[0][0] if texts else None
        for r in range(header + 1, len(data) + 1):
            k = key_of(data[r - 1][s.key_col - 1] if s.key_col - 1 < len(data[r - 1]) else None)
            if not k:
                continue
            s.all_rows.setdefault(match_key(k), []).append(r)
            if match_key(k) not in s.rows:
                s.rows[match_key(k)] = r
                s.keys[match_key(k)] = k
        years = Counter(y for c in s.value_cols for y in re.findall(r"20\d\d", str(_cell(data, header, c) or "")))
        s.year = max(years) if years else None
    return header, sides


def _cell(data, r, c):
    return data[r - 1][c - 1] if 0 < r <= len(data) and 0 < c <= len(data[r - 1]) else None


def _is_numbering(vals):
    """Столбец «№»: 1, 2, 3…"""
    ints = [int(v) for v in vals if isinstance(v, (int, float)) and float(v).is_integer()]
    return len(ints) >= 0.9 * len(vals) and ints[:5] == list(range(ints[0], ints[0] + min(5, len(ints)))) if ints else False


def _metric(h):
    """Показатель без годов и дат: «Бакия ба 01.01.2024 (30 фоиза)» → «бакия ба (30 фоиза)»."""
    t = norm_text(h)
    t = re.sub(r"\d{1,2}[./]\d{1,2}[./]\d{2,4}", " ", t)
    t = re.sub(r"(?<!\d)20\d\d(?!\d)", " ", t)
    return " ".join(t.split())


def _title(h):
    """Название показателя без даты и года: «Бакия ба 01.01.2024 (30 фоиза)» → «Бакия (30 фоиза)»."""
    t = str(h or "").replace("\n", " ")
    t = re.sub(r"\s*ба\s+\d{1,2}[./]\d{1,2}[./]\d{2,4}", "", t)
    t = re.sub(r"\d{1,2}[./]\d{1,2}[./]\d{2,4}|(?<!\d)20\d\d(?!\d)", "", t)
    return " ".join(t.split())


def compare(data):
    """→ (Diff, сторона1, сторона2) или None, если на листе нет двух списков с ИНН."""
    found = find_lists(data)
    if not found:
        return None
    header, (a, b) = found
    pairs = []
    used_b = set()
    for ca in a.value_cols:
        ma = _metric(_cell(data, header, ca))
        if not ma:
            continue
        for cb in b.value_cols:
            if cb not in used_b and _metric(_cell(data, header, cb)) == ma:
                pairs.append((_title(_cell(data, header, ca)), ca, cb))
                used_b.add(cb)
                break

    def name(side, k):
        return str(_cell(data, side.rows[k], side.name_col) or "").strip() if side.name_col else ""

    def values(side, k, cols):
        return {str(_cell(data, header, c) or get_column_letter(c)).replace("\n", " ").strip(): _cell(data, side.rows[k], c)
                for c in cols}

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
            v1, v2 = num(_cell(data, a.rows[k], ca)) or 0.0, num(_cell(data, b.rows[k], cb)) or 0.0
            vals.append((v1, v2))
            if abs(v2 - v1) > 0.005:
                changed.append((b.keys[k], name(b, k) or name(a, k), title, v1, v2, round(v2 - v1, 2)))
                diff_here = True
        (org_changed if diff_here else org_same).append((b.keys[k], name(b, k) or name(a, k), vals))
        same += not diff_here
    org_only1 = [(a.keys[k], name(a, k), [num(_cell(data, a.rows[k], p[1])) for p in pairs])
                 for k in a.rows if k not in b.rows]
    org_only2 = [(b.keys[k], name(b, k), [num(_cell(data, b.rows[k], p[2])) for p in pairs])
                 for k in b.rows if k not in a.rows]
    duplicates = [(i, side.keys[k], name(side, k), rows)
                  for i, side in ((1, a), (2, b)) for k, rows in side.all_rows.items() if len(rows) > 1]
    totals = [(title, round(sum(num(_cell(data, r, ca)) or 0.0 for r in a.rows.values()), 2),
               round(sum(num(_cell(data, r, cb)) or 0.0 for r in b.rows.values()), 2)) for title, ca, cb in pairs]
    return Diff(only1, only2, changed, same, len(common), names_differ, pairs, duplicates, totals,
                org_only1, org_only2, org_changed, org_same, org_names), a, b


def find_compares(sheets):
    """Все листы сравнения в книге (в порядке листов): название с ключевым словом и два списка с ИНН.
    → [(лист, Diff, список1, список2)]"""
    found = []
    for name, data in sheets.items():
        if compare_keyword(name):
            res = compare(data)
            if res:
                found.append((name, *res))
    return found


def best_compare(found):
    """Какой лист предложить первым: тот, где можно сравнить суммы."""
    return max(range(len(found)), key=lambda i: (len(found[i][1].pairs), i)) if found else None


def max_change(vals):
    """Наибольшая по модулю разница среди показателей: [(было, стало)] → число."""
    return max((abs((v2 or 0) - (v1 or 0)) for v1, v2 in vals), default=0.0)


def sorted_changes(diff):
    """Организации с изменившейся суммой — крупные разницы сверху."""
    return sorted(diff.org_changed, key=lambda o: -max_change(o[2]))
