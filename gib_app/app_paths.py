"""Папки программы в корне проекта: резервные копии и результаты глубокого анализа."""
import os
import sys

BACKUP_FOLDER = "_резервные копии"
ANALYSIS_FOLDER = "Свод по глубокому анализу"


def project_dir() -> str:
    """Корень проекта: рядом с gib_app при запуске из исходников, рядом с GIB.exe
    (а если exe лежит в dist — уровнем выше) при запуске программы."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.executable))
        if os.path.basename(base).lower() == "dist":
            base = os.path.dirname(base)
        return base
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def backup_dir() -> str:
    return os.path.join(project_dir(), BACKUP_FOLDER)


def analysis_dir() -> str:
    """Папка «Свод по глубокому анализу» (создаётся при первом обращении)."""
    folder = os.path.join(project_dir(), ANALYSIS_FOLDER)
    os.makedirs(folder, exist_ok=True)
    return folder
