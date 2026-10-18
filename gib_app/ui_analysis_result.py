"""Экран «Анализ готов»: коротко и по файлам — где сколько организаций и какая сумма, ошибки
(те же числа, что на листах ошибок в Excel)."""
import os
from html import escape

from analysis_model import part_summary
from report_errors import error_counts
from utils import format_money

MUTED = "#64748B"


def _int(n) -> str:
    return f"{n:,}".replace(",", " ")


def short_money(x) -> str:
    """10 814 763 244 → «10,81 млрд»; 345 200 → «345,2 тыс»."""
    if x is None:
        return "—"
    for div, word in ((1e9, "млрд"), (1e6, "млн"), (1e3, "тыс")):
        if abs(x) >= div:
            return f"{x / div:,.2f}".replace(",", " ").replace(".", ",") + f" {word}"
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


def _parts_table(data) -> str:
    rows = part_summary(data)
    th = "<td style='color:{c}' align='{a}'>{t}</td>"
    head = [("Часть", "left"), ("Файл", "left"), ("Организаций", "right"), ("Сумма", "right")]
    out = ["<tr>" + "".join(th.format(c=MUTED, a=a, t=t) for t, a in head) + "</tr>"]
    for r in rows:
        cells = [escape(r["part"].label), escape(r["part"].source), _int(r["orgs"]), short_money(r["sum"])]
        out.append("<tr>" + "".join(f"<td align='{a}'>{c}</td>" for c, (_, a) in zip(cells, head)) + "</tr>")
    return "<table cellspacing='0' cellpadding='5' width='100%'>" + "".join(out) + "</table>"


def _errors_table(data) -> str:
    """Виды ошибок × файлы — сколько строк на каждом листе ошибок и из какого файла."""
    counts = error_counts(data.details)
    if not counts:
        return "<b>Ошибки</b>: не найдено ✓"
    files = [f for f in dict.fromkeys(os.path.basename(p) for p in data.paths) if any(f in c for _, c, _ in counts)]
    th = "<td style='color:{c}' align='{a}'>{t}</td>"
    head = [("Ошибки", "left")] + [(escape(f), "right") for f in files] + [("Всего", "right")]
    out = ["<tr>" + "".join(th.format(c=MUTED, a=a, t=t) for t, a in head) + "</tr>"]
    for name, by_file, total in counts:
        cells = [escape(name)] + [_int(by_file[f]) if by_file.get(f) else "—" for f in files] + [f"<b>{_int(total)}</b>"]
        out.append("<tr>" + "".join(f"<td align='{a}'>{c}</td>" for c, (_, a) in zip(cells, head)) + "</tr>")
    return ("<table cellspacing='0' cellpadding='5' width='100%'>" + "".join(out) + "</table>"
            f"<span style='color:{MUTED}; font-size:10pt'>Подробно — в Excel, на листах с теми же названиями.</span>")


def result_view(res: dict, data) -> tuple:
    """→ (строки текста, метки [(заголовок, [(текст, цвет, подсказка)])], ошибки по файлам — под метками)."""
    lines = [f"<span style='color:{MUTED}'>{escape(os.path.basename(res['path']))} · папка «Свод по глубокому "
             f"анализу»</span>", "", _parts_table(data)]
    for c in data.checks:               # сверка с «Итого», которое пользователь указал в файле
        verdict = ("в ячейке не число" if c["diff"] is None else "<span style='color:#15803D'>сходится ✓</span>"
                   if abs(c["diff"]) < 0.01 else f"<b style='color:#B91C1C'>расхождение {format_money(c['diff'])}</b>")
        lines.append(f"Итого {escape(os.path.basename(c['file']))} ({c['cell']}): в файле "
                     f"{format_money(c['in_file']) if c['in_file'] is not None else escape(str(c['raw']))}, "
                     f"программа {format_money(c['program'])} — {verdict}")
    every = res.get("all_by_inn")
    if every:
        lines.append(f"<b>Одна орг. во всех файлах</b>: {_int(every['orgs'])} организаций, из них {_int(every['multi'])} есть в нескольких файлах "
                     f"({_int(every['rows'])} строк" + (", показаны не все)" if every["cut"] else ")")
                     + f" <span style='color:{MUTED}'>— смотрите лист «Одна орг. во всех файлах»</span>")
    recon = res.get("reconcile")
    if recon:
        lines.append(f"<b>Сверка со списком</b>: {_int(recon['orgs'])} организаций списка — их столбцы и платежи "
                     f"рядом <span style='color:{MUTED}'>(лист «Сверка со списком»)</span>")
    return lines, [], _errors_table(data)
