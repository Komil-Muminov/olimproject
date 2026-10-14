"""Экран «Анализ готов»: коротко и по файлам — где сколько организаций, кто появился и пропал, ошибки."""
import os
from html import escape

from analysis_model import FLAGS, STATUSES, part_summary
from report_analysis import COLORS, ERR_COLOR

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


def _chip(text: str, color: str) -> str:
    return f"<span style='background:#{color}'>&nbsp;{text}&nbsp;</span>"


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


def result_lines(res: dict, data) -> list:
    st, fl = res["statuses"], res["flags"]
    lines = [f"<span style='color:{MUTED}'>{escape(os.path.basename(res['path']))} · папка «Свод по глубокому "
             f"анализу»</span>", "", _parts_table(data)]
    if len(data.parts) > 1:
        lines += ["", f"<b>За весь период</b> <span style='color:{MUTED}'>(первая часть → последняя)</span>",
                  " ".join(_chip(f"{s} {_int(st[s])}", COLORS[s] if COLORS[s] != "FFFFFF" else "F1F5F9")
                           for s in STATUSES if st[s])]
    errors = [f for f in FLAGS if fl[f]]
    lines += ["", "<b>Ошибки</b>", " ".join(_chip(f"{f} {_int(fl[f])}", ERR_COLOR) for f in errors)
              if errors else "<span style='color:#15803D'>не найдено ✓</span>"]
    return lines
