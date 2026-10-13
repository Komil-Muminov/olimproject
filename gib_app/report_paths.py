"""Генерация путей к создаваемым отчетам."""
import os


def report_path(src: str, sheet: str) -> str:
    """«файл — различия (лист).xlsx» рядом с исходным; если уже есть — «… (2).xlsx»."""
    base = f"{os.path.splitext(src)[0]} — различия ({sheet})"
    path, n = base + ".xlsx", 2
    while os.path.exists(path):
        path, n = f"{base} ({n}).xlsx", n + 1
    return path


def check_report_path(src: str) -> str:
    """«файл — проверка ИНН.xlsx» рядом с исходным; если уже есть — «… (2).xlsx»."""
    base = f"{os.path.splitext(src)[0]} — проверка ИНН"
    path, n = base + ".xlsx", 2
    while os.path.exists(path):
        path, n = f"{base} ({n}).xlsx", n + 1
    return path
