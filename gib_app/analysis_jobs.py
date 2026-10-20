"""Фоновая задача глубокого анализа после чтения файлов: собрать таблицу с выбранной разбивкой и парой
«было → стало» и записать Excel «Анализ»."""
from analysis_scan import build
from report_analysis import write_analysis_report


def make_report(path, raw, mode, pair, mappings=None, threshold=None, progress=lambda p, t: None) -> dict:
    """mappings — колонки из карточек файлов; threshold — порог суммы (от, до) для листа «Суммы …»."""
    progress(5, "Собираю таблицу…")
    data = build(*raw, mode=mode, pair=pair)
    res = write_analysis_report(path, data, progress, mappings, threshold)
    res["data"] = data
    return res

