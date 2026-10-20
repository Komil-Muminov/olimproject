"""Лист «Суммы от …»: все записи из всех файлов, где сумма строки (по колонкам «Сумма») в пределах порога.
Остальной анализ не меняется — порог только выделяет крупные (или нужные) записи."""
from report_rows import collect, in_range, row_sum, threshold_text, write_rows_sheet


def write_threshold_sheet(wb, paths: list, mappings: dict, threshold) -> dict | None:
    """Только если порог задан. → {"text", "rows", "total", "cut"}."""
    if not threshold or not mappings:
        return None
    found = collect(paths, mappings, lambda values, m: bool(m.get("sums")) and in_range(row_sum(values, m), threshold))
    text = threshold_text(threshold)
    rows, total = sum(x[1] for x in found["per"]), round(sum(x[2] or 0 for x in found["per"]), 2)
    write_rows_sheet(wb, f"Суммы {text}"[:31], f"СУММЫ {text.upper()} — {rows:,} записей".replace(",", " "),
                     "Записи из всех файлов, где сумма строки (по колонкам «Сумма») в пределах порога. Файлы без "
                     "колонки «Сумма» не проверяются.", found)
    return {"text": text, "rows": rows, "total": total, "cut": found["cut"]}
