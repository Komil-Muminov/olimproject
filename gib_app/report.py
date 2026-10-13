"""Генерация Excel-отчетов (фасад для вызова специализированных генераторов)."""

from report_check import write_check_report
from report_compare import write_compare_report
from report_paths import check_report_path, report_path

__all__ = [
    "report_path",
    "check_report_path",
    "write_compare_report",
    "write_check_report",
]
