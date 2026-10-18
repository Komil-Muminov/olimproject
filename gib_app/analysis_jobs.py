"""Фоновая задача глубокого анализа после чтения файлов: собрать таблицу и записать Excel."""
from analysis_scan import build
from report_analysis import write_analysis_report


def make_report(path, raw, mappings=None, progress=lambda p, t: None) -> dict:
    """mappings — что выбрал пользователь в окне «Листы и ИНН» (для листа «Как считали»)."""
    progress(5, "Собираю таблицу…")
    data = build(*raw)
    res = write_analysis_report(path, data, progress, mappings)
    res["data"] = data
    return res
