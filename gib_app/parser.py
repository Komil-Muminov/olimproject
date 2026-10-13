"""Разбор назначения платежа: поиск всех «ГИБ XXXX-сумма».

Правила те же, что у макроса в Excel (assets/modGIB.bas.txt, ParseGib/ReadAmount),
чтобы программа и встроенная автоматика понимали назначения одинаково.
ИНН, ГВБ, номера заявок и прочие числа без слова «ГИБ» перед ними пропускаются.
"""
import re
from dataclasses import dataclass, field

# «ГИБ» (или GIB), до 4 разделителей, код (3–8 цифр; если больше 4 — ГИБ первые 4), затем тире/двоеточие/«=» или пробел
GIB_RE = re.compile(r"(?:ГИБ|GIB)[ :.\-#№]{0,4}(\d{3,8})(?!\d)( *[-:=] *| +)(?=\d)", re.IGNORECASE)
SPACES = dict.fromkeys(map(ord, "\t\r\n           "
                                " ​  　﻿"), " ")
DASHES = dict.fromkeys(map(ord, "­‐‑‒–—―−﹣－"), "-")


@dataclass
class ParsedRow:
    src_row: int                 # номер строки в выписке
    text: str                    # назначение
    gibs: dict = field(default_factory=dict)  # код ГИБ -> сумма
    statement_amount: float | None = None     # сумма из выписки (для сверки)
    include: bool = True
    done: bool = False                        # суммы ГИБ в этой строке уже стоят

    @property
    def gib_total(self) -> float:
        return round(sum(self.gibs.values()), 2)

    @property
    def status(self) -> str:
        if self.done:
            return "Уже заполнено"
        if not self.gibs:
            return "Без ГИБ"
        if self.statement_amount is not None and abs(self.gib_total - self.statement_amount) > 0.005:
            return "Не сходится"
        if len(self.gibs) > 1:
            return "Несколько ГИБ"
        return "OK"


def read_amount(t: str, p: int):
    """Сумма с позиции p (там цифра) → (сумма, позиция после неё).

    Тысячи: пробел/точка/запятая перед группой ровно из 3 цифр («1 500,50», «1.500,50»),
    только если до неё не больше 3 цифр. Копейки: 1–2 цифры после запятой или точки.
    Поэтому «ГИБ 2211-140 020012345» — это 140, а не 140 020.
    """
    m = re.match(r"\d+", t[p:])
    whole, q = m.group(), p + m.end()
    first_len, frac = len(whole), ""
    while q < len(t) and t[q] in " .,":
        run = re.match(r"\d*", t[q + 1:]).group()
        if t[q] in ".," and len(run) in (1, 2):
            frac, q = run, q + 1 + len(run)
            break
        if len(run) == 3 and first_len <= 3:
            whole, q = whole + run, q + 4
            continue
        break
    return round(float(whole + "." + (frac or "0")), 2), q


def parse_purpose_ex(text: str):
    """→ ({код_ГИБ: сумма}, [длинные коды]). Один и тот же ГИБ в строке суммируется.

    Код — 3–8 цифр, как в макросе образца. Если цифр больше 4 (например, 22110 — смета
    с подпунктом), ГИБ — первые 4 (2211), а длинный код возвращается отдельно: строку
    подсвечиваем жёлтым для проверки.
    """
    t = str(text or "").translate(SPACES).translate(DASHES)
    result, long_codes, pos = {}, [], 0
    while m := GIB_RE.search(t, pos):
        raw = m.group(1)
        code = raw[:4] if len(raw) > 4 else raw
        if len(raw) > 4:
            long_codes.append(raw)
        amount, pos = read_amount(t, m.end())
        result[code] = round(result.get(code, 0.0) + amount, 2)
    return result, long_codes


def parse_purpose(text: str) -> dict:
    """Возвращает {код_ГИБ: сумма}."""
    return parse_purpose_ex(text)[0]


def normalize(text) -> str:
    """Текст назначения для сравнения: без лишних пробелов и регистра."""
    return " ".join(str(text or "").split()).casefold()


def to_number(value) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(" ", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None
