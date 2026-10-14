"""Свод реестров платежей: суммы по ИНН за каждый файл (период), двойные платежи, неправильные ИНН,
странные даты. Чистая логика без Excel: строки подаются по одной через add()."""
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta

INN_RE = re.compile(r"^\d{9}$")
MONTHS = ("Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек")
MAIN_SHARE = 0.01          # месяц, где меньше 1% строк файла, — «чужой» (странная дата)
KEEP_ROWS = 2000           # подробности храним, пока месяц маленький
ODD_GAP_DAYS = 31          # дата документа дальше месяца от периода файла — «странная»


@dataclass
class Party:
    """Плательщик или получатель: суммы и количество платежей по файлам, варианты названия."""
    inn: str
    names: Counter = field(default_factory=Counter)
    sums: dict = field(default_factory=lambda: defaultdict(float))
    counts: Counter = field(default_factory=Counter)

    @property
    def name(self) -> str:
        return self.names.most_common(1)[0][0] if self.names else ""

    @property
    def total(self) -> float:
        return round(sum(self.sums.values()), 2)


@dataclass
class FileInfo:
    path: str
    sheet: str
    rows: int = 0
    total: float = 0.0
    months: Counter = field(default_factory=Counter)
    month_sums: dict = field(default_factory=lambda: defaultdict(float))
    label: str = ""
    main_months: list = field(default_factory=list)


class Registry:
    def __init__(self):
        self.files: list[FileInfo] = []
        self.payers: dict[str, Party] = {}
        self.recs: dict[str, Party] = {}
        self.dups = []              # (файл, Payment, (файл, строка) первого такого же)
        self.bad_inn = []           # (файл, Payment, «плательщик»/«получатель», ИНН)
        self.odd_dates = []         # (файл, Payment) — заполняется в finish()
        self.order = []             # файлы по времени (заполняется в finish)
        self._seen = {}
        self._month_rows = []

    def start_file(self, path: str, sheet: str) -> int:
        """Новый источник: файл (или лист файла, если реестров в нём несколько)."""
        self.files.append(FileInfo(path, sheet))
        self._month_rows.append(defaultdict(list))
        return len(self.files) - 1

    def _party(self, book, inn, name, f, amount):
        p = book.get(inn)
        if p is None:
            p = book[inn] = Party(inn)
        p.names[" ".join(name.upper().split())] += 1
        p.sums[f] += amount
        p.counts[f] += 1

    def add(self, f: int, p):
        info = self.files[f]
        info.rows += 1
        info.total += p.amount
        for book, side, inn, name in ((self.payers, "плательщик", p.payer_inn, p.payer),
                                      (self.recs, "получатель", p.rec_inn, p.rec)):
            if inn is None:             # в документе нет этой стороны
                continue
            self._party(book, inn, name, f, p.amount)
            if not INN_RE.match(inn):
                self.bad_inn.append((f, p, side, inn))
        # двойной платёж: совпадает всё, включая основание (банк платит за разных людей одинаковыми суммами
        # с короткими номерами — их основания разные, это не повтор)
        key = hash((p.doc, p.date, p.payer_inn, p.rec_inn, p.amount, " ".join(p.purpose.lower().split())))
        first = self._seen.get(key)
        if first is None:
            self._seen[key] = (f, p.row)
        else:
            self.dups.append((f, p, first))
        if p.date:
            m = (p.date.year, p.date.month)
            info.months[m] += 1
            info.month_sums[m] += p.amount
            kept = self._month_rows[f].get(m, [])
            if kept is not None:
                kept.append(p)
                self._month_rows[f][m] = kept if len(kept) <= KEEP_ROWS else None

    def finish(self):
        """После всех файлов: периоды файлов (по основным месяцам) и странные даты."""
        self._seen.clear()
        for f, info in enumerate(self.files):
            main = info.main_months = sorted(m for m, c in info.months.items() if c >= info.rows * MAIN_SHARE)
            info.label = period_label(main)
            if main:
                lo = date(*main[0], 1) - timedelta(days=ODD_GAP_DAYS)
                y, m = main[-1]
                hi = date(y + m // 12, m % 12 + 1, 1) + timedelta(days=ODD_GAP_DAYS - 1)
            for m, kept in self._month_rows[f].items():
                # платёжка от 31.12 в январском файле — обычное дело; странно, если дата дальше месяца
                if m not in main and kept and main:
                    self.odd_dates += [(f, p) for p in kept if not lo <= p.date <= hi]
        self._month_rows = []
        # порядок файлов — по времени, чтобы «пропал / появился» считались правильно
        self.order = sorted(range(len(self.files)), key=lambda f: self.files[f].main_months[:1] or [(9999, 0)])


def period_label(months: list) -> str:
    """[(2024, 1) … (2024, 4)] → «Янв–Апр 2024»."""
    if not months:
        return "без дат"
    (y1, m1), (y2, m2) = months[0], months[-1]
    if (y1, m1) == (y2, m2):
        return f"{MONTHS[m1 - 1]} {y1}"
    if y1 == y2:
        return f"{MONTHS[m1 - 1]}–{MONTHS[m2 - 1]} {y1}"
    return f"{MONTHS[m1 - 1]} {y1} – {MONTHS[m2 - 1]} {y2}"
