"""Список организаций, где ИНН не нашёлся сам: пользователь указывает столбцы (ИНН, название, сумма),
таблица читается по ним — даже если ИНН записаны с ошибками или строк мало."""
import re

from compare_finder import key_in_text, match_key
from compare_models import SideList
from registry_columns import guess_header
from utils import cell_val, find_col, is_total_label, norm_text

AS_LIST = "list"        # пользователь сказал: читать файл как обычный список организаций
RAW = "raw"             # лист без ИНН: в свод как есть, без сопоставления организаций
NAME_KEY = "≈"          # ключ организации без ИНН: «≈» + название без кавычек и форм
SPLIT_RE = re.compile(r"[,;\n]+")      # несколько организаций в одной ячейке
FORMS = {"ооо", "оао", "зао", "пао", "ао", "ип", "тоо", "чдмм", "чск", "ксп", "квд", "ддтт", "мдт", "ммд", "ч"}
FIELDS = ("inn", "name", "amount")
TITLES = {"inn": "ИНН", "name": "Название", "amount": "Сумма"}
FIND = {"inn": ("инн", "рма", "рмя", "inn", "tin"),
        "name": ("наимен", "название", "контрагент", "организац", "номи", "ташкилот", "name"),
        "amount": ("сумма", "маблаг", "amount")}


def as_list(mapping) -> bool:
    """Файл читать как список (а не как реестр платежей)?"""
    return mapping == AS_LIST or isinstance(mapping, dict) and mapping.get("kind") == AS_LIST


def list_ok(cols: dict) -> bool:
    return cols.get("inn") is not None


def _find(values) -> dict:
    heads = [norm_text(v) for v in values]
    found = {k: find_col(heads, p) for k, p in FIND.items()}
    return {k: (c - 1 if c else None) for k, c in found.items()}


def list_guess(sheets: dict) -> dict | None:
    """Для окна выбора столбцов: самый заполненный лист и его шапка. Пустой файл → None."""
    filled = {name: data for name, data in sheets.items() if any(v not in (None, "") for r in data for v in r)}
    if not filled:
        return None
    name = max(filled, key=lambda n: sum(any(v not in (None, "") for v in r) for r in filled[n]))
    return dict(guess_header(filled[name], _find), sheet=name)


def side_guess(data: list, sheet: str, s) -> dict:
    """Для окна выбора столбцов — по найденной таблице: ИНН и название уже выбраны, сумму выберет пользователь."""
    row = lambda r: [str(v).strip() if v is not None else "" for v in (data[r - 1] if r <= len(data) else [])]
    return {"sheet": sheet, "head": s.header - 1, "heads": row(s.header), "sample": row(s.header + 1),
            "cols": {"inn": s.key_col - 1, "name": s.name_col - 1 if s.name_col else None, "amount": None}}


def _key(v) -> str | None:
    """ИНН из ячейки, указанной пользователем: цифры (без пробелов), иначе текст как есть."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if not s or is_total_label(s):
        return None
    digits = key_in_text(s) or re.sub(r"\s+", "", s)
    return digits if digits.isdigit() else s


def org_key(name) -> str:
    """Название для сопоставления без ИНН: без кавычек, знаков и форм (ООО, ҶСК, ЧДММ…)."""
    words = re.sub(r"[^0-9a-zа-яё]+", " ", norm_text(name)).split()
    return "".join(w for w in words if w not in FORMS)


def split_cell(v) -> list:
    """Ячейка с несколькими значениями («ООО А, ООО Б» / «123…; 456…») → части; иначе — одна."""
    if isinstance(v, str) and SPLIT_RE.search(v.strip()):
        return [x.strip() for x in SPLIT_RE.split(v) if x.strip()]
    return [v]


def transpose(data: list) -> list:
    """«Перевёрнутая» таблица (организации по столбцам) → обычная (организации по строкам)."""
    width = max((len(r) for r in data), default=0)
    return [[r[i] if i < len(r) else None for r in data] for i in range(width)]


def _row_keys(data, r: int, s) -> list:
    """→ [(ключ, ИНН как в файле, название или None)] — организации в строке r (одна или несколько в ячейке)."""
    parts = split_cell(cell_val(data, r, s.key_col))
    names = split_cell(cell_val(data, r, s.name_col)) if s.name_col and not s.by_name else []
    out = []
    for i, v in enumerate(parts):
        if s.by_name:
            text = str(v or "").strip()
            if text and not is_total_label(text) and org_key(text):
                out.append((NAME_KEY + org_key(text), "", text))
        elif (inn := _key(v)):
            out.append((match_key(inn) or inn, inn, names[i] if len(names) == len(parts) > 1 else None))
    return out


def mapped_table(data: list, mapping: dict) -> SideList:
    """Таблица по столбцам пользователя: {"sheet", "head", "cols": {"inn", "name", "amount"}, "sums": [столбцы]}
    (с 0; "row0" / "col0" — сдвиг, если номера считаны от первой заполненной ячейки листа)."""
    r0, c0 = mapping.get("row0", 0), mapping.get("col0", 0)
    c = {k: (v + c0 + 1 if v is not None else None) for k, v in mapping["cols"].items() if k != "amount"}
    money = mapping.get("sums", []) + ([mapping["cols"]["amount"]] if mapping["cols"].get("amount") is not None else [])
    header = r0 + mapping["head"] + 1
    cols = sorted(v for v in c.values() if v)
    by_name = bool(mapping.get("by_name"))
    s = SideList(c["name"] if by_name else c["inn"], c.get("name"), cols, [v + c0 + 1 for v in money], block=cols,
                 header=header, fixed_money=True, by_name=by_name, separate=bool(mapping.get("separate")),
                 repeats=bool(mapping.get("repeats")))
    last = min(len(data), mapping.get("last") or len(data))       # «C5-C300» — дальше 300-й строки не читать
    for r in range(header + 1, last + 1):
        keys = _row_keys(data, r, s)
        for k, inn, name in keys:
            s.all_rows.setdefault(k, []).append(r)
            s.keys.setdefault(k, inn)
            if len(keys) > 1:
                s.shared.add(k)
            if name:
                s.row_names[k, r] = name
            s.last_row = r
    return s
