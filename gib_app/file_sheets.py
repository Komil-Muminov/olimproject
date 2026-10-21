"""Листы выбранных файлов для карточек фильтра: первые строки (чтобы найти начало данных и подписи колонок),
скрыт ли лист и где формулы без значений. Колонки пользователь вписывает сам — программа их не угадывает."""
from excel_sheet_info import empty_formulas, workbook_sheets
from file_format import real_path

PREVIEW_ROWS = 300          # столько первых строк листа читается для карточки
FULL_READ_LIMIT = 200_000   # листы больше — без поиска формул (долго)


def _sheet_rows(sheet) -> list:
    rows = []
    for i, values in enumerate(sheet.iter_rows()):
        rows.append(["" if v is None else v for v in values])
        if i + 1 >= PREVIEW_ROWS:
            break
    return rows


def read_sheets(path: str) -> dict:
    """→ {"path", "sheets": {лист: {"rows", "col0", "height", "hidden", "formulas"}}} (пустые листы пропущены)."""
    from python_calamine import CalamineWorkbook

    wb, info = CalamineWorkbook.from_path(real_path(path)), workbook_sheets(path)
    sheets = {}
    for name in wb.sheet_names:
        sheet = wb.get_sheet_by_name(name)
        rows = _sheet_rows(sheet)
        if any(v != "" for r in rows for v in r):
            # iter_rows: строки — с первой строки листа, столбцы — с первого заполненного (col0)
            col0 = (sheet.start or (0, 0))[1]
            formulas = empty_formulas(path, name) if sheet.height <= FULL_READ_LIMIT else {}
            sheets[name] = {"rows": rows, "col0": col0, "height": sheet.height,
                            "hidden": info.get(name, {}).get("hidden", False),
                            "formulas": {c - col0: n for c, n in formulas.items()}}
    return {"path": path, "sheets": sheets}


def read_files(paths: list, progress=lambda p, t: None) -> list:
    out = []
    for k, path in enumerate(paths):
        progress(int(100 * k / len(paths)), f"Смотрю файл {k + 1} из {len(paths)}…")
        out.append(read_sheets(path))
    return out
