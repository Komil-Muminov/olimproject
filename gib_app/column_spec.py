"""Колонки, вписанные пользователем (как в ClearAudit / Olimv3):
  C        — вся колонка C;       C5 — с 5-й строки до конца листа;
  C5-300 / C5-C300 — строки с 5 по 300;   B, C / B5-B20, C5-C20 — несколько колонок."""
import re

# русские буквы, похожие на латинские (набрали в русской раскладке): С → C, А → A…
LOOKALIKE = str.maketrans("АВСЕНКМОРТХавсенкмортх", "ABCEHKMOPTXABCEHKMOPTX")
PART_RE = re.compile(r"^([A-Z]{1,3})(\d+)?(?:-([A-Z]{1,3})?(\d+))?$")


class SpecError(ValueError):
    """Непонятная запись колонки — текст ошибки показывается пользователю."""


def col_index(letters: str) -> int:
    """A → 0, Z → 25, AA → 26."""
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def parse(text: str) -> list:
    """«B5-B20, C» → [(столбец с 0, первая строка или None, последняя строка или None), …]; пусто → []."""
    out = []
    for raw in str(text or "").translate(LOOKALIKE).upper().replace(" ", "").replace(";", ",").split(","):
        if not raw:
            continue
        m = PART_RE.match(raw)
        if not m:
            raise SpecError(f"«{raw}» — непонятно. Пишите, например: C, C5 или C5-C300")
        col, first, col2, last = m.groups()
        if col2 and col2 != col:
            raise SpecError(f"«{raw}» — диапазон должен быть в одной колонке (например, C5-C300)")
        first, last = (int(first) if first else None), (int(last) if last else None)
        if first and last and last < first:
            raise SpecError(f"«{raw}» — конец диапазона раньше начала")
        out.append((col_index(col), first, last))
    return out


def one_column(text: str, title: str):
    """Поле, где нужна одна колонка (ИНН, Название, Дата…). → (столбец, первая строка, последняя) или None."""
    parts = parse(text)
    if len(parts) > 1:
        raise SpecError(f"{title}: нужна одна колонка, а указано {len(parts)}")
    return parts[0] if parts else None
