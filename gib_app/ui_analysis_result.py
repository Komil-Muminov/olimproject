"""Экран «Анализ готов»: коротко и по файлам — где сколько организаций, кто появился и пропал, ошибки
(те же числа, что на листах ошибок в Excel)."""
import os
from html import escape

from analysis_model import STATUSES, part_summary, status_hints
from report_colors import COLORS
from report_errors import error_counts

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


def _delta(n, sign: str, color: str) -> str:
    """+18 432 зелёным / −21 389 красным; нет изменений или первая часть — прочерк."""
    return "—" if not n else f"<span style='color:{color}'>{sign}{_int(n)}</span>"


def _parts_table(data) -> str:
    rows = part_summary(data)
    compare = len(rows) > 1
    th = "<td style='color:{c}' align='{a}'>{t}</td>"
    head = [("Часть", "left"), ("Файл", "left"), ("Организаций", "right"), ("Сумма", "right")]
    head += [("Появились", "right"), ("Пропали", "right")] if compare else []
    out = ["<tr>" + "".join(th.format(c=MUTED, a=a, t=t) for t, a in head) + "</tr>"]
    for r in rows:
        cells = [escape(r["part"].label), escape(r["part"].source), _int(r["orgs"]), short_money(r["sum"])]
        if compare:
            cells += [_delta(r["new"], "+", "#15803D"), _delta(r["gone"], "−", "#B91C1C")]
        out.append("<tr>" + "".join(f"<td align='{a}'>{c}</td>" for c, (_, a) in zip(cells, head)) + "</tr>")
    return "<table cellspacing='0' cellpadding='5' width='100%'>" + "".join(out) + "</table>"


def _part_name(part) -> str:
    """«M1-4.xlsx (Янв–Апр 2024)»; если подпись и есть файл — один раз."""
    return part.source if part.label in part.source else f"{part.source} ({part.label})"


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
    st = res["statuses"]
    lines = [f"<span style='color:{MUTED}'>{escape(os.path.basename(res['path']))} · папка «Свод по глубокому "
             f"анализу»</span>", "", _parts_table(data)]
    recon = res.get("reconcile")
    if recon:
        lines.append(f"<b>Сверка со списком</b>: {_int(recon['orgs'])} организаций списка — их столбцы и платежи "
                     f"рядом <span style='color:{MUTED}'>(лист «Сверка со списком»)</span>")
    chips = []
    if len(data.parts) > 1:
        lines.append(f"<span style='color:{MUTED}; font-size:10pt'>Появились / Пропали — по сравнению "
                     f"с предыдущим файлом. Наведите на метку ниже — подсказка.</span>")
        hints = status_hints(data.parts, data.compare)
        a, b = data.compare
        chips.append((f"Сравнение: {_part_name(data.parts[a])} → {_part_name(data.parts[b])}",
                      [(f"{s}: {_int(st[s])}", COLORS[s] if COLORS[s] != "FFFFFF" else "F1F5F9", hints[s])
                       for s in STATUSES if st[s]]))
    return lines, chips, _errors_table(data)
