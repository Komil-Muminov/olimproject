"""ИНН в ячейке: отдельным числом или внутри текста; ключ сравнения — без нулей впереди."""
import re

from utils import norm_text

KEY_RE = re.compile(r"^\d{8,14}$")
TEXT_KEY_RE = re.compile(r"[\s,;(\[]*(?:инн|рма|рмя|inn)\D{0,5}?(\d{8,14})(?!\d)[)\]]?", re.I)


def key_of(v) -> str | None:
    """ИНН/РМА: 8–14 цифр (ведущие нули могли потеряться, если Excel хранит число)."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s if KEY_RE.match(s) else None


def key_in_text(v) -> str | None:
    """ИНН внутри текста: «ООО А (ИНН 123456789)» → 123456789."""
    m = TEXT_KEY_RE.search(norm_text(v)) if isinstance(v, str) else None
    return m.group(1) if m else None


def match_key(k: str) -> str:
    """Ключ сопоставления без ведущих нулей."""
    return k.lstrip("0")
