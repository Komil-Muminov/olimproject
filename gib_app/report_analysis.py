"""«В Excel» из глубокого анализа: то, что сейчас видно в таблице, и подробности выбранных ошибок."""
from datetime import datetime

from openpyxl import Workbook

from analysis_model import BAD_INN, DUP_PAY, DUP_ROW, NAMES, ODD_DATE, PAYER, entity_key
from report_fast import table_sheet

DETAILS = {   # (флаг, вид данных) → (лист, цвет, пояснение, шапка, ширины, столбцы сумм, столбцы дат)
    (DUP_PAY, "payments"): ("Двойные платежи", "B91C1C", "Платёж повторяется: одинаковые №, дата, оба ИНН и сумма.",
                            ["№", "Файл", "Строка", "№ док.", "Дата", "Плательщик", "ИНН плат.", "Получатель",
                             "ИНН получ.", "Сумма", "Первый раз: файл", "строка"],
                            [7, 14, 9, 13, 12, 36, 12, 36, 12, 16, 14, 9], (10,), (5,)),
    (BAD_INN, "payments"): ("Неправильные ИНН", "EA580C", "ИНН не из 9 цифр. Справа — ИНН, вписанный в название.",
                            ["№", "Файл", "Строка", "Чей ИНН", "Что стоит", "Название", "ИНН из названия", "№ док.",
                             "Дата", "Сумма"], [7, 14, 9, 13, 14, 46, 14, 13, 12, 16], (10,), (9,)),
    (ODD_DATE, "payments"): ("Странные даты", "64748B", "Дата документа не из периода файла — возможно, опечатка.",
                             ["№", "Файл", "Период", "Строка", "№ док.", "Дата документа", "Плательщик",
                              "Получатель", "Сумма"], [7, 14, 15, 9, 13, 14, 40, 40, 16], (9,), (6,)),
    (DUP_ROW, "list"): ("Повторы в таблице", "7C3AED", "Один ИНН записан в таблице несколько раз.",
                        ["№", "Файл", "Таблица", "ИНН", "Название", "Строки", "Сколько раз"],
                        [7, 16, 40, 13, 46, 18, 12], (), ()),
    (NAMES, "list"): ("Разные названия", "0891B2", "У одного ИНН название записано по-разному.",
                      ["№", "Файл", "Таблица", "ИНН", "Как записано", "Строка"], [7, 16, 40, 13, 56, 9], (), ()),
    (BAD_INN, "list"): ("Неправильные ИНН (списки)", "EA580C", "ИНН выглядит неправильно; справа — как должно быть.",
                        ["№", "Файл", "Таблица", "Строка", "Что стоит", "Название", "Проблема", "Как должно быть"],
                        [7, 16, 40, 9, 16, 46, 36, 16], (), ()),
}


def write_analysis_report(path, data, rows, sel, flags, progress=lambda p, t: None) -> dict:
    """rows — видимые строки [(организация, что случилось, всего)]; sel — выбранные части;
    flags — выбранные фильтры ошибок (по ним добавляются листы с подробностями)."""
    labels = [data.parts[i].label for i in sel]
    with_role = len(data.roles) > 1
    head = ["№", "ИНН", "Название"] + (["Кто"] if with_role else []) + labels + ["Всего", "Что случилось", "Ошибки"]
    first = 4 + with_role
    progress(20, "Пишу таблицу…")
    wb = Workbook(write_only=True)
    table_sheet(wb, "Таблица", "ГЛУБОКИЙ АНАЛИЗ", "0F172A",
                f"Создано {datetime.now():%d.%m.%Y %H:%M} · то, что было видно в программе после фильтров.",
                head, [7, 13, 46] + [12] * with_role + [17] * len(labels) + [18, 16, 30],
                [[n, e.inn, e.name] + ([e.role] if with_role else []) + [e.sums[i] if e.counts[i] else None for i in sel]
                 + [total, st, ", ".join(sorted(e.flags))] for n, (e, st, total) in enumerate(rows, 1)],
                money=range(first, first + len(labels) + 1))
    shown = {entity_key(e) for e, _, _ in rows}
    known = {entity_key(e) for e in data.entities}
    sheets = 1
    for (flag, kind), (name, color, hint, dhead, widths, money, dates) in DETAILS.items():
        if flag not in flags:
            continue
        seen, found = set(), []
        for key, r in data.details.get((flag, kind), []):       # один платёж — у плательщика и получателя
            if (key in shown or key not in known) and id(r) not in seen:
                seen.add(id(r))
                found.append(r)
        if found:
            progress(40 + 8 * sheets, f"Пишу «{name}»…")
            table_sheet(wb, name, name.upper(), color, hint, dhead, widths,
                        [[n] + r for n, r in enumerate(found, 1)], money=money, dates=dates)
            sheets += 1
    if NAMES in flags and data.registry:
        _registry_names(wb, data.registry, [e for e, _, _ in rows if NAMES in e.flags and e.role != "организация"])
    progress(95, "Сохраняю…")
    wb.save(path)
    return {"path": path, "rows": len(rows)}


def _registry_names(wb, reg, ents):
    """Варианты названий у ИНН из реестров (для видимых строк с «разные названия»)."""
    out = []
    for e in ents:
        party = (reg.payers if e.role == PAYER else reg.recs).get(e.inn)
        if party:
            out += [[e.inn, e.role, name, cnt] for name, cnt in party.names.most_common()]
    table_sheet(wb, "Названия в реестрах", "РАЗНЫЕ НАЗВАНИЯ (РЕЕСТРЫ)", "0891B2",
                "Как записан один ИНН в разных платежах и сколько раз.", ["№", "ИНН", "Кто", "Как записано", "Раз"],
                [7, 13, 12, 60, 9], [[n] + r for n, r in enumerate(out, 1)])
