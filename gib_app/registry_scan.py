"""Чтение нескольких реестров в один свод (для фонового потока) и выгрузка платежей одного ИНН."""
import os

from excel_com import WriteError
from registry_read import read_registry
from registry_stats import Registry

ASK_SCORE = 2       # один файл: столько узнанных заголовков — уже похоже на реестр, спросить про столбцы


def _read(path, mappings, single):
    try:
        return read_registry(path, (mappings or {}).get(path))
    except Exception as e:
        if single:          # один обычный файл — откроется по-старому (Глубокий анализ по ИНН)
            return [], None
        raise WriteError(f"Не удалось прочитать «{os.path.basename(path)}»:\n{e}")


def scan_registries(paths, mappings=None, progress=lambda p, t: None) -> dict:
    """Все файлы (все листы-реестры в них) → один Registry.
    Файл не узнан: {"needs_mapping": путь, "guess": догадка} — пусть пользователь укажет столбцы;
    один файл, совсем не похожий на реестр: {"not_registry": путь}."""
    reg = Registry()
    n = len(paths)
    for k, path in enumerate(paths):
        name = os.path.basename(path)
        base = int(100 * k / n)
        progress(base, f"Открываю {name} ({k + 1} из {n})…")
        sheets, guess = _read(path, mappings, n == 1)
        if not sheets:
            if guess and (n > 1 or guess["score"] >= ASK_SCORE):
                return {"needs_mapping": path, "guess": guess}
            return {"not_registry": path}
        for sheet, total, rows in sheets:
            f = reg.start_file(path, sheet)
            for i, p in enumerate(rows):
                reg.add(f, p)
                if i % 50000 == 0:
                    progress(base + int(100 / n * i / max(total, 1)), f"Считаю {name} ({k + 1} из {n})…")
    progress(99, "Подвожу итоги…")
    reg.finish()
    return {"registry": reg}


def payments_of(paths, inn: str, mappings=None, progress=lambda p, t: None) -> list:
    """Все платежи, где ИНН — плательщик или получатель: [(файл, Payment)]."""
    out = []
    for k, path in enumerate(paths):
        progress(int(100 * k / len(paths)), f"Ищу в {os.path.basename(path)}…")
        sheets, _ = read_registry(path, (mappings or {}).get(path))
        for _, _, rows in sheets:
            out += [(path, p) for p in rows if inn in (p.payer_inn, p.rec_inn)]
    return out
