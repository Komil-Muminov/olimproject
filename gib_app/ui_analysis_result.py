"""Экран «Анализ готов»: коротко и по файлам — где сколько организаций, кто появился и пропал, ошибки."""
import os
from html import escape

from analysis_model import FLAG_HINTS, FLAGS, STATUSES, part_summary, status_hints
from report_colors import COLORS, ERR_COLOR

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


def result_view(res: dict, data) -> tuple:
    """→ (строки текста, метки [(заголовок, [(текст, цвет, подсказка)])]) для экрана «Анализ готов»."""
    st, fl = res["statuses"], res["flags"]
    lines = [f"<span style='color:{MUTED}'>{escape(os.path.basename(res['path']))} · папка «Свод по глубокому "
             f"анализу»</span>", "", _parts_table(data)]
    chips = []
    if len(data.parts) > 1:
        lines.append(f"<span style='color:{MUTED}; font-size:10pt'>Появились / Пропали — по сравнению "
                     f"с предыдущим файлом. Наведите на метку ниже — подсказка.</span>")
        hints = status_hints(data.parts, data.compare)
        a, b = data.compare
        chips.append((f"Что случилось: «{data.parts[a].label}» → «{data.parts[b].label}»",
                      [(f"{s}: {_int(st[s])}", COLORS[s] if COLORS[s] != "FFFFFF" else "F1F5F9", hints[s])
                       for s in STATUSES if st[s]]))
    errors = [f for f in FLAGS if fl[f]]
    chips.append(("Ошибки", [(f"{f[:1].upper() + f[1:]}: {_int(fl[f])}", ERR_COLOR,
                              FLAG_HINTS[f] + " Подробно — на листе с ошибками в Excel.") for f in errors]
                  or [("Не найдено ✓", "DCFCE7", "Ни одной ошибки в выбранных файлах.")]))
    return lines, chips
