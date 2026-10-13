"""Карточка найденного ИНН: название, таблица «Период | Платил | Получал», выгрузка в Excel."""
from html import escape

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QScrollArea, QVBoxLayout

from utils import format_money


def build_found_panel(win):
    """Карточка под полем ввода; скрыта, пока ничего не искали."""
    box = win.reg_found_box = QFrame()
    box.setObjectName("foundBox")
    lay = QVBoxLayout(box)
    lay.setContentsMargins(16, 10, 16, 10)
    lay.setSpacing(6)
    win.reg_found = QLabel()
    win.reg_found.setTextFormat(Qt.RichText)
    win.reg_found.setWordWrap(True)
    win.reg_found.setTextInteractionFlags(Qt.TextSelectableByMouse)
    win.reg_found.setAlignment(Qt.AlignTop | Qt.AlignLeft)
    scroll = QScrollArea()          # на маленьком экране таблица прокручивается, окно не растёт
    scroll.setWidget(win.reg_found)
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    scroll.setMinimumHeight(90)
    scroll.viewport().setStyleSheet("background: white;")
    win.reg_found.setStyleSheet("background: white;")
    lay.addWidget(scroll, 1)
    btns = QHBoxLayout()
    win.reg_found_note = QLabel()
    win.reg_found_note.setObjectName("muted")
    btns.addWidget(win.reg_found_note, 1)
    win.reg_export_btn = QPushButton("📄  Все платежи этого ИНН — в Excel")
    win.reg_export_btn.setObjectName("secondary")
    win.reg_export_btn.setVisible(False)
    btns.addWidget(win.reg_export_btn)
    lay.addLayout(btns)
    box.setVisible(False)
    return box


def _money(r, i):
    return format_money(r.sums[i]) if r and r.counts[i] else "—"


def found_html(res: dict, labels: list) -> str:
    """HTML карточки: res = {"payer": PartyRow|None, "rec": PartyRow|None} (см. registry_analysis.search)."""
    p, r = res["payer"], res["rec"]
    main = p or r
    other = r.name if p and r and r.name != p.name else ""
    head = (f"<div style='font-size:13pt; font-weight:600; color:#0F172A'>{escape(main.name)}</div>"
            f"<div style='color:#64748B'>ИНН {main.inn}" + (f" · как получатель: {escape(other)}" if other else "") + "</div>")
    th = "<td bgcolor='#F1F5F9' align='{a}'><b>{t}</b></td>"
    rows = ["<tr>" + th.format(a="left", t="Период") + th.format(a="right", t="Платил")
            + th.format(a="right", t="Получал") + "</tr>"]
    for i, lab in enumerate(labels):
        rows.append(f"<tr><td>{lab}</td><td align='right'>{_money(p, i)}</td><td align='right'>{_money(r, i)}</td></tr>")
    total = lambda x: format_money(x.total) if x else "—"
    count = lambda x: str(sum(x.counts)) if x else "—"
    rows.append(f"<tr><td bgcolor='#EFF6FF'><b>Всего</b></td><td bgcolor='#EFF6FF' align='right'><b>{total(p)}</b></td>"
                f"<td bgcolor='#EFF6FF' align='right'><b>{total(r)}</b></td></tr>")
    rows.append(f"<tr><td style='color:#64748B'>Платежей</td><td align='right' style='color:#64748B'>{count(p)}</td>"
                f"<td align='right' style='color:#64748B'>{count(r)}</td></tr>")
    table = "<table width='100%' cellspacing='0' cellpadding='4'>" + "".join(rows) + "</table>"
    return head + table


def show_found(win, res: dict, labels: list, raw: str):
    hit = res["payer"] or res["rec"]
    win.reg_found_inn = hit.inn if hit else None
    win.reg_found.setText(found_html(res, labels) if hit else
                          f"<p style='font-size:12pt'>«{escape(raw)}» не найден ни среди плательщиков, ни среди получателей.</p>")
    win.reg_found_note.setText(" · ".join(f"{t}: {x.status}" for t, x in (("Платил", res["payer"]), ("Получал", res["rec"]))
                                          if x and x.status))
    win.reg_export_btn.setVisible(bool(hit))
    win.reg_found_box.setVisible(True)
