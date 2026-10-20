"""Колонки, вписанные от руки в окне «Как читать файл» → настройка чтения листа (mapping для scan_files).
Строки берутся из поля ИНН: «C5-C300» — с 5 по 300; «C» — с первой строки, где в колонке C стоит ИНН."""
from openpyxl.utils import get_column_letter

from analysis_list_map import AS_LIST
from analysis_propose import REGISTRY
from column_spec import SpecError, one_column, parse
from compare_finder import key_in_text, key_of
from utils import is_total_label, to_number

MAIN = (("inn", "ИНН"), ("name", "Название"), ("sums", "Сумма"), ("total", "Итого (колонка)"))
REGISTRY_FIELDS = (("rec_inn", "ИНН получателя"), ("rec", "Получатель"), ("date", "Дата"), ("doc", "№ документа"),
                   ("purpose", "Основание"))
TITLES = dict(MAIN + REGISTRY_FIELDS)


def _first_data_row(rows: list, col: int, by_name: bool) -> int:
    """Номер (с 0) первой строки, где в колонке ИНН (или название, если ИНН нет) — данные."""
    for i, r in enumerate(rows):
        v = r[col] if col < len(r) else ""
        if (by_name and i > 0 and str(v).strip() and str(rows[i - 1][col] if col < len(rows[i - 1]) else "").strip()) \
                or (not by_name and (key_of(v) or key_in_text(v))):
            return i
    letter = get_column_letter(col + 1)
    raise SpecError(f"В колонке {letter} в первых {len(rows)} строках нет ИНН — укажите, с какой строки начинать, "
                    f"например {letter}5")


def labels_row(rows: list, head: int) -> int:
    """Строка с названиями колонок (для подписей): ближайшая над данными, где хотя бы два текста и это не
    «Итого / Ҳамагӣ». Нет такой — строка прямо над данными."""
    for j in range(head, max(head - 6, -1), -1):
        texts = [v for v in rows[j] if isinstance(v, str) and v.strip() and to_number(v) is None]
        if len(texts) >= 2 and not is_total_label(texts[0]):           # строка итогов — «Итого» первым словом
            return j
    return head


def _labels(sheet: dict, head: int, typed: dict) -> dict:
    """Для «Как считали»: поле → что вписано и как называются эти колонки в файле («C5-C300 → C: РМА»)."""
    heads = sheet["rows"][head] if 0 <= head < len(sheet["rows"]) else []

    def name(c: int) -> str:
        i = c - sheet["col0"]
        text = str(heads[i]).strip() if 0 <= i < len(heads) else ""
        return f"{get_column_letter(c + 1)}: {text}" if text else get_column_letter(c + 1)

    out = {}
    for key, title in TITLES.items():
        if typed.get(key) and key != "total":
            out[title] = f"{typed[key]}  →  " + "; ".join(name(c) for c, _, _ in parse(typed[key]))
        elif typed.get(key):
            out[title] = typed[key]
    return out


def to_mapping(sheet_name: str, sheet: dict, typed: dict, opts: dict) -> dict:
    """typed — что вписано в поля ({"inn": "C5-C300", "sums": "G, H", …}); opts — галочки и отбор.
    Ошибка ввода → SpecError с понятным текстом."""
    col0 = sheet["col0"]

    def rel(c: int, title: str) -> int:
        if c < col0:
            raise SpecError(f"{title}: колонка {get_column_letter(c + 1)} на этом листе пустая")
        return c - col0

    one = {key: one_column(typed.get(key), title) for key, title in TITLES.items() if key not in ("sums", "total")}
    total = one_column(typed.get("total"), "Итого")             # колонка; строки — из поля ИНН
    by_name = bool(opts.get("by_name"))
    anchor = one["inn"] or (one["name"] if by_name else None)
    if not anchor:
        raise SpecError("Укажите колонку ИНН (например, C или C5-C300)")
    registry = bool(one["rec_inn"] or one["rec"])
    sums = [rel(c, "Сумма") for c, _, _ in parse(typed.get("sums"))]
    if registry and not sums:
        raise SpecError("Для реестра (есть получатель) укажите колонку суммы")
    if registry and not one["inn"]:
        raise SpecError("Для реестра нужна колонка ИНН плательщика")
    col, first, last = anchor
    head = first - 2 if first else _first_data_row(sheet["rows"], rel(col, "ИНН"), by_name and not one["inn"]) - 1
    cols = {key: (rel(v[0], TITLES[key]) if v else None) for key, v in one.items()}
    if registry:
        cols = {"payer_inn": cols["inn"], "payer": cols["name"], "rec_inn": cols["rec_inn"], "rec": cols["rec"],
                "date": cols["date"], "doc": cols["doc"], "purpose": cols["purpose"], "amount": sums}
    else:
        cols = {"inn": cols["inn"], "name": cols["name"]}
    m = {"sheet": sheet_name, "head": head, "last": last, "row0": 0, "col0": col0, "kind": REGISTRY if registry else AS_LIST,
         "cols": cols, "typed": {TITLES[k]: v for k, v in typed.items() if v}, "total_col": rel(total[0], "Итого") if total else None,
         "hidden_sheet": bool(sheet.get("hidden"))}
    m["labels_row"] = labels_row(sheet["rows"], head) if head >= 0 else head
    m["labels"] = _labels(sheet, m["labels_row"], typed)
    m.update({k: v for k, v in opts.items() if v})
    if not registry:
        m["sums"] = sums
    return m
