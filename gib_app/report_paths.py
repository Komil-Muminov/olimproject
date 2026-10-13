"""Генерация путей к создаваемым отчетам (рядом с исходным файлом, без перезаписи)."""
import os


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


def registry_report_path(src: str, period: str) -> str:
    """«Свод по ИНН (Янв–Дек 2024).xlsx» в папке первого реестра."""
    return unique_path(os.path.join(os.path.dirname(src), f"Свод по ИНН ({period})"))


def inn_report_path(src: str, inn: str) -> str:
    """«Платежи ИНН 040000011.xlsx» в папке первого реестра."""
    return unique_path(os.path.join(os.path.dirname(src), f"Платежи ИНН {inn}"))
