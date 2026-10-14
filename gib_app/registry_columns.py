"""Столбцы реестра платежей по заголовкам (русский, таджикский, английский) и догадка о шапке,
когда реестр не распознан (тогда пользователь сам укажет столбцы)."""
from utils import find_col, norm_text

# norm_text: строчные, таджикские буквы → русские (маблағ → маблаг, ҳуҷҷат → хуччат)
COLS = {
    "doc": ("№ док", "номер док", "№док", "№ хуччат", "раками хуччат", "doc"),
    "date": ("дата док", "дата плат", "сана", "date"),
    "payer": ("плательщик", "наименование плат", "пардохткунанда", "номи пардохткунанда", "payer"),
    "payer_inn": ("инн плат", "инн отправ", "рма плат", "рмя плат", "рма пардохт", "payer inn", "inn payer"),
    "amount": ("сумма", "маблаг", "amount"),
    "rec": ("получатель", "наименование получ", "гиранда", "номи гиранда", "beneficiary", "receiver", "recipient"),
    "rec_inn": ("инн получ", "рма получ", "рмя получ", "рма гиранда", "receiver inn", "inn receiver",
                "beneficiary inn"),
    "purpose": ("основание", "назначение", "асос", "максад", "purpose"),
}
TITLES = {"payer": "Плательщик (название)", "payer_inn": "ИНН плательщика", "rec": "Получатель (название)",
          "rec_inn": "ИНН получателя", "amount": "Сумма", "doc": "№ документа", "date": "Дата документа",
          "purpose": "Основание"}
HEAD_ROWS = 20


def is_enough(cols: dict) -> bool:
    """Хватает для свода: сумма и хотя бы один ИНН (плательщика или получателя)."""
    return cols.get("amount") is not None and (cols.get("payer_inn") is not None or cols.get("rec_inn") is not None)


def find_columns(head) -> dict:
    """Найденные по заголовкам столбцы (0-based); ненайденные — None."""
    heads = [norm_text(v) for v in head]
    found = {k: find_col(heads, p) for k, p in COLS.items()}
    return {k: (c - 1 if c else None) for k, c in found.items()}


def registry_columns(head) -> dict | None:
    """Столбцы реестра или None, если по этой строке реестр не узнаётся."""
    cols = find_columns(head)
    return cols if is_enough(cols) else None


def _text(v) -> str:
    return "" if v is None else str(v).strip()


def two_row_columns(top, bottom) -> dict | None:
    """Шапка в две строки: «Плательщик» (объединённая ячейка) над «Наименование | ИНН» → заголовки
    «ИНН Плательщик», «Наименование Плательщик»… Верхний текст тянется вправо над подписями нижней строки."""
    pairs, last = [], ""
    for i in range(max(len(top), len(bottom))):
        t, b = _text(top[i] if i < len(top) else None), _text(bottom[i] if i < len(bottom) else None)
        last = t or (last if b else "")
        pairs.append((b, last))
    for first, second in ((0, 1), (1, 0)):          # «ИНН Плательщик» или «Плательщик ИНН»
        cols = registry_columns([" ".join(x for x in (p[first], p[second]) if x) for p in pairs])
        if cols:
            return cols
    return None


def guess_header(rows: list, finder=None) -> dict:
    """Самая похожая на шапку строка среди первых: больше всего узнанных заголовков, иначе — больше
    всего текстовых ячеек. → {"head": номер строки, "heads": [текст], "sample": [первая строка данных],
    "cols": {…}, "score": узнано}. finder — свой поиск столбцов (по умолчанию — столбцы реестра)."""
    finder = finder or find_columns
    best = {"head": 0, "heads": [], "cols": {}, "score": -1, "texts": -1}
    for i, values in enumerate(rows[:HEAD_ROWS]):
        cols = finder(values)
        score = sum(c is not None for c in cols.values())
        texts = sum(isinstance(v, str) and bool(v.strip()) for v in values)
        if (score, texts) > (best["score"], best["texts"]):
            sample = rows[i + 1] if i + 1 < len(rows) else []
            best = {"head": i, "heads": [str(v).strip() if v is not None else "" for v in values],
                    "sample": [str(v).strip() if v is not None else "" for v in sample],
                    "cols": cols, "score": score, "texts": texts}
    return best
