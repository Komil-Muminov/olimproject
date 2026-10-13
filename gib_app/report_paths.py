"""Пути к создаваемым отчетам (без перезаписи): сравнение и проверка — рядом с исходным файлом,
свод реестров и выгрузки по ИНН — в папке «Свод по глубокому анализу»."""
import os

from app_paths import analysis_dir


def unique_path(base: str) -> str:
    """«base.xlsx»; если уже есть — «base (2).xlsx», «base (3).xlsx»…"""
    path, n = base + ".xlsx", 2
    while os.path.exists(path):
        path, n = f"{base} ({n}).xlsx", n + 1
    return path


def report_path(src: str, sheet: str) -> str:
    """«файл — различия (лист).xlsx» рядом с исходным."""
    return unique_path(f"{os.path.splitext(src)[0]} — различия ({sheet})")


def check_report_path(src: str) -> str:
    """«файл — проверка ИНН.xlsx» рядом с исходным."""
    return unique_path(f"{os.path.splitext(src)[0]} — проверка ИНН")


def registry_report_path(period: str) -> str:
    """«Свод по ИНН (Янв–Дек 2024).xlsx» в папке «Свод по глубокому анализу» (корень проекта)."""
    return unique_path(os.path.join(analysis_dir(), f"Свод по ИНН ({period})"))


def inn_report_path(inn: str) -> str:
    """«Платежи ИНН 040000011.xlsx» в папке «Свод по глубокому анализу»."""
    return unique_path(os.path.join(analysis_dir(), f"Платежи ИНН {inn}"))
