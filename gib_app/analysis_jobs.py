"""Фоновые задачи глубокого анализа после чтения файлов: собрать таблицу с выбранной разбивкой и парой
«было → стало» и записать Excel («Анализ» или «Сводная»)."""
from analysis_scan import build
from report_analysis import write_analysis_report
from report_pivot import write_pivot_report


def make_report(path, raw, mode, pair, progress=lambda p, t: None) -> dict:
    progress(5, "Собираю таблицу…")
    data = build(*raw, mode=mode, pair=pair)
    res = write_analysis_report(path, data, progress)
    res["data"] = data
    return res


def make_pivot(path, raw, mode, pair, progress=lambda p, t: None) -> dict:
    progress(5, "Собираю таблицу…")
    return write_pivot_report(path, build(*raw, mode=mode, pair=pair), progress)
