"""ИНН: правильный ли (9 цифр) и что с ним не так — простыми словами."""
import re

INN_DIGITS = 9


def inn_ok(inn: str) -> bool:
    """ИНН правильный, если в нём ровно 9 цифр (пробелы и знаки ошибкой не считаются)."""
    return len(re.sub(r"\D", "", inn or "")) == INN_DIGITS


def inn_problem(inn: str) -> str:
    """Что не так с ИНН — простыми словами."""
    digits = re.sub(r"\D", "", inn or "")
    if not inn:
        return "ИНН пустой"
    if digits and not digits.strip("0"):
        return "вместо ИНН стоит 0 (не указан)"
    signs = " (есть буквы или знаки)" if digits != inn else ""
    tail = f" — возможно, потерян 0 в начале (должно быть {digits.zfill(INN_DIGITS)})" if len(digits) == 8 else ""
    return f"{len(digits)} цифр вместо {INN_DIGITS}{signs}{tail}"
