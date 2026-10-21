"""Пути к создаваемым отчетам (без перезаписи) — в папке «Свод по глубокому анализу»."""
import os

from app_paths import analysis_dir


def unique_path(base: str) -> str:
    """«base.xlsx»; если уже есть — «base (2).xlsx», «base (3).xlsx»…"""
    path, n = base + ".xlsx", 2
    while os.path.exists(path):
        path, n = f"{base} ({n}).xlsx", n + 1
    return path


def analysis_report_path() -> str:
    """«Анализ 14.10.2026 10-30.xlsx» в папке «Свод по глубокому анализу»."""
    from datetime import datetime
    return unique_path(os.path.join(analysis_dir(), f"Анализ {datetime.now():%d.%m.%Y %H-%M}"))


def filter_report_path(title: str) -> str:
    """«ИНН 040058284, суммы от 2 000,00 21.10.2026 09-30.xlsx» в папке «Свод по глубокому анализу»."""
    from datetime import datetime
    safe = "".join(ch for ch in title if ch not in r'\/:*?"<>|')         # знаки, недопустимые в имени файла
    return unique_path(os.path.join(analysis_dir(), f"{safe} {datetime.now():%d.%m.%Y %H-%M}"))
