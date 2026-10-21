"""Общие вспомогательные функции для работы с числами, текстом и ячейками."""
from constants import SUM_ROW_LABELS, TJ_TO_RU

TOTAL_WORDS = {" ".join(w.translate(TJ_TO_RU).lower().split()) for w in SUM_ROW_LABELS}


def to_number(value) -> float | None:
    """Преобразует значение в float или None, если число некорректно."""
    if value is None or value == "" or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(" ", "").replace(" ", "").replace(",", ".")
    sign = -1.0 if s.startswith("(") and s.endswith(")") else 1.0     # бухгалтерский минус: (1 000,00)
    try:
        return sign * float(s.strip("()"))
    except ValueError:
        return None


def is_total_label(v) -> bool:
    """Строка итогов: первое слово — целиком «Итого», «Всего», «Ҷамъ», «Ҳамагӣ»… («Ҷамъияти …» — это ООО, не итог)."""
    words = norm_text(v).replace(":", " ").replace(".", " ").split()
    return bool(words) and words[0] in TOTAL_WORDS


def to_num_safe(v) -> float:
    """Возвращает число с плавающей точкой или 0.0."""
    res = to_number(v)
    return res if res is not None else 0.0


def norm_text(v) -> str:
    """Нормализует текст: нижний регистр, единые пробелы, таджикские буквы в русские."""
    return " ".join(str(v or "").translate(TJ_TO_RU).lower().split())


def find_col(heads: list, prefixes: tuple, after: int = 0) -> int | None:
    """Номер столбца (с 1), чей нормализованный заголовок начинается с одного из prefixes."""
    return next((c for c, h in enumerate(heads, 1) if c > after and h.startswith(prefixes)), None)


def format_money(x) -> str:
    """Форматирует число в денежный вид: 1 234,56."""
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")
