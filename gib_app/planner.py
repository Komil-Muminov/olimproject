"""Анализ книги Excel и подготовка плана действий."""
from detect import by_header, detect, sheet_score
from excel_reader import Book
from parser import parse_purpose_ex
from utils import to_number


class Plan:
    """Что нашлось в файле и что будет сделано."""

    def __init__(self, path: str):
        self.path = path
        book = Book(path)

        scores = {n: sheet_score(d) for n, d in book.sheets.items()}
        self.sheet = max(scores, key=scores.get) if scores else None
        data = book.sheets.get(self.sheet, [])
        self.layout = detect(data) if self.sheet else None
        self.items, self.no_gib, self.done, self.mismatch = [], [], [], []
        self.long_codes = []
        self.codes, self.new_codes, self.total = [], [], 0.0
        self.header_only = None

        if not self.layout:
            for name in [self.sheet] + [n for n in book.sheets if n != self.sheet]:
                lay = by_header(book.sheets.get(name, []))
                if lay:
                    self.sheet, self.header_only = name, lay
                    break
            return

        lay = self.layout
        for r in range(lay.first_row, lay.last_row + 1):
            v = Book.value(data, r, lay.purpose_col)
            if not isinstance(v, str) or not v.strip():
                continue
            gibs, longs = parse_purpose_ex(v)
            if longs:
                self.long_codes.append((r, longs))
            if not gibs:
                self.no_gib.append(r)
                continue
            if lay.gib_cols and all(code in lay.gib_cols and to_number(Book.value(data, r, lay.gib_cols[code])) == a
                                    for code, a in gibs.items()):
                self.done.append(r)
                continue
            if lay.amount_col:
                paid = to_number(Book.value(data, r, lay.amount_col))
                if paid is not None and abs(paid - sum(gibs.values())) > 0.005:
                    self.mismatch.append(r)
            self.items.append((r, gibs))

        self.codes = sorted({c for _, g in self.items for c in g})
        self.new_codes = [c for c in self.codes if c not in lay.gib_cols]
        self.total = round(sum(sum(g.values()) for _, g in self.items), 2)
