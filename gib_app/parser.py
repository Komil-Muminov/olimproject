"""Разбор назначения платежа: поиск всех «ГИБ XXXX-сумма»."""
import re
from dataclasses import dataclass, field

# ГИБ 2211-140,00 | ГИБ:2211 - 1 200,50 | гиб 2211–140
GIB_RE = re.compile(
    r"ГИБ[\s:№.]*(\d{4})\s*[-–—]\s*"
    r"(\d{1,3}(?:[  ]\d{3})+|\d+)"   # целая часть, допускаем пробел-разделитель тысяч
    r"(?:[.,](\d{1,2}))?",                # копейки
    re.IGNORECASE,
)


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


def parse_purpose(text: str) -> dict:
    """Возвращает {код_ГИБ: сумма}. Один и тот же ГИБ в строке суммируется."""
    result = {}
    for code, whole, frac in GIB_RE.findall(text or ""):
        amount = float(whole.replace(" ", "").replace(" ", "") + "." + (frac or "0"))
        result[code] = round(result.get(code, 0.0) + amount, 2)
    return result


def normalize(text) -> str:
    """Текст назначения для сравнения: без лишних пробелов и регистра."""
    return " ".join(str(text or "").split()).casefold()


def to_number(value) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(" ", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None
