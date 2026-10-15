"""Для окна «Как читать файл»: предупреждение о формулах без сохранённого результата в колонках суммы."""
from openpyxl.utils import get_column_letter


def formula_warning(name: str, sheet: dict, st: dict) -> str:
    """Колонки суммы (с 0, от первой заполненной колонки листа) с формулами без результата — суммы выйдут пустыми."""
    bad = [f"{get_column_letter(sheet['col0'] + c + 1)} ({n})" for c, n in sheet.get("formulas", {}).items()
           if c in st["sums"]]
    return (f"«{name}»: в столбцах {', '.join(bad)} формулы без сохранённого результата — суммы будут пустыми. "
            "Откройте файл в Excel и сохраните его.") if bad else ""
