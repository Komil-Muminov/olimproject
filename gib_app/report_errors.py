"""Листы ошибок: «Ошибки в реестрах» и «Ошибки в таблицах» (только если есть). Ошибки по группам —
у каждой группы цветной заголовок (наведите — подсказка), пояснение и одинаковый порядок столбцов."""
from openpyxl.cell import WriteOnlyCell
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from report_fast import DATE, title_rows
from report_styles import BORDER, CENTER, NUM, fill_color

PAY_HEAD = ["Файл", "Строка в файле", "№ док.", "Дата", "Сумма"]
# группа: (заголовок, цвет, что это и как проверить, шапка)
GROUPS = {
    "dup": ("ДВОЙНОЙ ПЛАТЁЖ", "B91C1C",
            "Каждая пара строк — один и тот же платёж: «платёж» — где он записан впервые, «повтор» — где ещё раз "
            "(совпадают №, дата, плательщик, получатель, сумма и основание). Проверьте, не провели ли его дважды.",
            PAY_HEAD + ["Что это", "Плательщик", "ИНН плательщика", "Получатель", "ИНН получателя",
                        "Основание"]),
    "bad_pay": ("НЕПРАВИЛЬНЫЙ ИНН", "EA580C",
                "ИНН должен быть из 9 цифр. «ИНН в названии» — если в названии организации вписан ИНН, скорее "
                "всего правильный именно он.",
                PAY_HEAD + ["Чей ИНН", "Организация", "Что стоит", "Что не так", "ИНН в названии"]),
    "odd": ("СТРАННАЯ ДАТА", "64748B",
            "Дата платёжки больше чем на месяц раньше или позже периода файла (например, 2004 или 2029 год). "
            "Возможна опечатка в дате.",
            PAY_HEAD + ["Период файла", "Плательщик", "ИНН плательщика", "Получатель", "ИНН получателя"]),
    "dup_row": ("ПОВТОР ИНН В ТАБЛИЦЕ", "7C3AED",
                "Один ИНН записан в одной таблице несколько раз — возможно, организация задвоена.",
                ["Файл", "Таблица", "Строки", "ИНН", "Организация", "Сколько раз"]),
    "same_name": ("ОДНО НАЗВАНИЕ — РАЗНЫЕ ИНН", "0891B2",
                  "Организации с одинаковым названием, но разными ИНН (каждая группа — одно название). Возможна "
                  "ошибка в ИНН; бывает и законно — например, филиалы.",
                  ["Название", "Кто", "ИНН", "Всего", "Записей"]),
    "bad_list": ("НЕПРАВИЛЬНЫЙ ИНН", "EA580C",
                 "ИНН в таблице выглядит неправильно. «Как должно быть» — если программа может подсказать.",
                 ["Файл", "Таблица", "Строка", "Что стоит", "Организация", "Что не так", "Как должно быть"]),
}
# лист: (группы, ширины столбцов с «№», столбцы сумм, столбцы дат — с 1, без «№»)
SHEETS = {
    "Ошибки в реестрах": (("dup", "bad_pay", "odd"), [6, 13, 10, 16, 12, 16, 16, 34, 16, 34, 16, 50], (5,), (4,)),
    "Ошибки в таблицах": (("dup_row", "bad_list"), [6, 16, 40, 16, 16, 44, 36, 16], (), ()),
    "Одно название — разные ИНН": (("same_name",), [6, 56, 13, 13, 18, 10], (4,), ()),
}
BIGGEST_FIRST = ("dup", "bad_pay")      # крупные суммы сверху
PAIRED = ("dup", "same_name")           # строки идут группами: первый столбец — номер группы
PAIR_SHADE = "F1F5F9"


def _cell(ws, v, **style):
    c = WriteOnlyCell(ws, value=v)
    for k, val in style.items():
        setattr(c, k, val)
    return c


def write_errors_sheets(wb, details: dict) -> list:
    """→ названия созданных листов (если ошибок нет — один лист «Ошибки» с «не найдено»)."""
    made = []
    for name, (groups, widths, money, dates) in SHEETS.items():
        if any(details.get(g) for g in groups):
            _sheet(wb, name, [(g, details[g]) for g in groups if details.get(g)], widths, money, dates)
            made.append(name)
    if not made:
        ws = wb.create_sheet("Ошибки")
        title_rows(ws, "ОШИБКИ", "16A34A", "Ошибок не найдено ✓", 4)
        made.append("Ошибки")
    return made


def _sheet(wb, name, groups, widths, money, dates):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = "C2410C"
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    total = sum(len({r[0] for r in rows}) if g in PAIRED else len(rows) for g, rows in groups)   # группы — по одной
    title_rows(ws, f"{name.upper()} — {total}", "C2410C",
               "Ошибки по группам. В каждой группе: что это, как проверить и строки, где найдено.", len(widths))
    for g, rows in groups:
        title, color, hint, head = GROUPS[g]
        paired = g in PAIRED
        if g in BIGGEST_FIRST:
            rows = sorted(rows, key=lambda r: (-(r[4 + paired] or 0), r[0] if paired else 0))
        count = len({r[0] for r in rows}) if paired else len(rows)
        band = fill_color(color)
        t = _cell(ws, f"{title} — {count}", font=Font(bold=True, size=13, color="FFFFFF"), fill=band,
                  alignment=Alignment(vertical="center", indent=1))
        t.comment = Comment(hint, "Распределение по ГИБ", width=320, height=90)
        ws.append([t] + [_cell(ws, None, fill=band) for _ in head])
        ws.append([_cell(ws, hint, font=Font(italic=True, color="475569"))])
        ws.append([_cell(ws, h, font=Font(bold=True), fill=fill_color("E2E8F0"), border=BORDER, alignment=CENTER)
                   for h in ["№"] + head])
        pair_no = {}
        for k, r in enumerate(rows, 1):
            if paired:                      # № — номер пары по порядку, пары через одну с фоном
                k = pair_no.setdefault(r[0], len(pair_no) + 1)
                r = r[1:]
            shade = {"fill": fill_color(PAIR_SHADE)} if paired and k % 2 == 0 else {}
            ws.append([_cell(ws, k, **shade)] + [
                _cell(ws, v, number_format=NUM, **shade) if i in money and v is not None
                else _cell(ws, v, number_format=DATE, **shade) if i in dates and v is not None
                else _cell(ws, v, **shade) for i, v in enumerate(r, 1)])
        ws.append([])
