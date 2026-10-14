"""«Заполнить ГИБ»: если в выписке или во втором файле столбцы не узнаны — спросить пользователя."""
from PySide6.QtWidgets import QDialog

from detect import Layout
from gib_columns import (
    STATEMENT_FIELDS, STATEMENT_TITLES, TARGET_FIELDS, TARGET_TITLES, book_guess, statement_cols, statement_ok,
    target_cols, target_ok
)
from ui_column_dialog import ColumnDialog


def ask_statement_columns(win, path: str) -> dict | None:
    """Столбцы выписки от пользователя: {"sheet", "head", "cols"} или None (отмена)."""
    guess = book_guess(path, statement_cols)
    if not guess:
        return None
    dlg = ColumnDialog(win, path, guess, STATEMENT_FIELDS, STATEMENT_TITLES, statement_ok,
                       need="Обязательно: <b>Назначение</b> и <b>Сумма расхода</b>.",
                       warn="Нужно указать назначение платежа и сумму расхода.", required=("purpose", "amount"))
    return dlg.mapping() if dlg.exec() == QDialog.Accepted else None


def ask_target_columns(win, path: str):
    """Куда заполнять, от пользователя: (лист, Layout, столбцы) или None (отмена)."""
    guess = book_guess(path, target_cols)
    if not guess:
        return None
    dlg = ColumnDialog(win, path, guess, TARGET_FIELDS, TARGET_TITLES, target_ok,
                       need="Обязательно: столбец <b>назначений</b>. Суммы пойдут в «Расходы» — если такого "
                            "столбца нет, программа добавит «Ҳамагӣ» и «Расходы» сама.",
                       warn="Нужно указать столбец назначений.", required=("purpose",))
    if dlg.exec() != QDialog.Accepted:
        return None
    m = dlg.mapping()
    head = m["head"] + 1                     # строки в Excel — с 1
    return m["sheet"], Layout(m["cols"]["purpose"] + 1, head + 1, head + 1, head), m["cols"]
