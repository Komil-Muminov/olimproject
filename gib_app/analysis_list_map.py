"""Список организаций, где ИНН не нашёлся сам: пользователь указывает столбцы (ИНН, название, сумма),
таблица читается по ним — даже если ИНН записаны с ошибками или строк мало."""
import re

from compare_finder import key_in_text, match_key
from compare_models import SideList
from constants import SUM_ROW_LABELS
from registry_columns import guess_header
from utils import cell_val, find_col, norm_text

AS_LIST = "list"        # пользователь сказал: читать файл как обычный список организаций
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
    if not s or norm_text(s).startswith(SUM_ROW_LABELS):
        return None
    digits = key_in_text(s) or re.sub(r"\s+", "", s)
    return digits if digits.isdigit() else s


def mapped_table(data: list, mapping: dict) -> SideList:
    """Таблица по столбцам пользователя: {"sheet", "head" (0-based), "cols": {"inn", "name", "amount"}}."""
    c = {k: (v + 1 if v is not None else None) for k, v in mapping["cols"].items()}
    header = mapping["head"] + 1
    cols = sorted(v for v in c.values() if v)
    s = SideList(c["inn"], c.get("name"), cols, [c["amount"]] if c.get("amount") else [], block=cols, header=header)
    for r in range(header + 1, len(data) + 1):
        k = _key(cell_val(data, r, s.key_col))
        if k:
            s.all_rows.setdefault(match_key(k) or k, []).append(r)
            s.keys.setdefault(match_key(k) or k, k)
            s.last_row = r
    return s
