"""«Итого» в каждой строке: колонку указал пользователь. Сумма указанных колонок «Сумма» в строке должна быть
равна её «Итого»; не сходится — строка выписывается, организация подсвечивается (при наведении — подробности)."""
from utils import cell_val, format_money, to_num_safe, to_number

TOLERANCE = 0.01        # копейка — округление, не расхождение


def row_total_diffs(data, s, money: list) -> list:
    """→ [(ключ организации, строка, итого в файле, сумма колонок)] — где не сходится (пустое «Итого» не в счёт)."""
    if not s.total_col:
        return []
    out = []
    for k, rows in s.all_rows.items():
        for r in rows:
            total = to_number(cell_val(data, r, s.total_col))
            if total is None:
                continue
            got = round(sum(to_num_safe(cell_val(data, r, c)) for c in money), 2)
            if abs(got - total) >= TOLERANCE:
                out.append((k, r, total, got))
    return out


def diff_text(fname: str, row: int, total: float, got: float) -> str:
    """Подсказка при наведении: где и на сколько не сходится."""
    return (f"{fname}, строка {row}: «Итого» в файле {format_money(total)}, сумма колонок {format_money(got)}, "
            f"разница {format_money(got - total)}")
