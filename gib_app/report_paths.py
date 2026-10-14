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


def pivot_report_path() -> str:
    """«Сводная 14.10.2026 10-30.xlsx» в папке «Свод по глубокому анализу»."""
    from datetime import datetime
    return unique_path(os.path.join(analysis_dir(), f"Сводная {datetime.now():%d.%m.%Y %H-%M}"))


def inn_report_path(inn: str) -> str:
    """«Платежи ИНН 040000011.xlsx» в папке «Свод по глубокому анализу»."""
    return unique_path(os.path.join(analysis_dir(), f"Платежи ИНН {inn}"))
