"""Таблица стилей (QSS) интерфейса приложения."""

STYLE = """
* { font-family: 'Segoe UI'; font-size: 10pt; }
QMainWindow, QWidget#root { background: #F5F7FB; }
QLabel#appTitle { font-size: 22pt; font-weight: 700; color: #0F172A; }
QLabel#muted { color: #64748B; }
QLabel#fileName { font-size: 13pt; font-weight: 600; color: #0F172A; }
QLabel#bigText { font-size: 12pt; color: #0F172A; line-height: 150%; }
QLabel#warn { font-size: 11pt; color: #B45309; }
QLabel#resultOk { font-size: 20pt; font-weight: 700; color: #16A34A; }
QLabel#resultBad { font-size: 20pt; font-weight: 700; color: #DC2626; }
QFrame#card { background: white; border: 1px solid #E2E8F0; border-radius: 16px; }
QPushButton { background: #F1F5F9; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 16px; color: #0F172A; }
QPushButton:hover { background: #E2E8F0; }
QPushButton#dropZone { border: 2px dashed #CBD5E1; border-radius: 16px; color: #64748B; font-size: 15pt; background: white; }
QPushButton#dropZone:hover { border-color: #2563EB; color: #2563EB; background: #EFF6FF; }
QPushButton#primary { background: #2563EB; color: white; border: none; padding: 12px 28px; font-weight: 600; font-size: 12pt; }
QPushButton#primary:hover { background: #1D4ED8; }
QPushButton#primary:disabled { background: #94A3B8; }
QFrame#cmpBox { background: transparent; border: none; }
QLabel#cmpTitle { font-size: 11pt; font-weight: 600; color: #334155; padding-bottom: 4px; }
QLabel#cmpInfo { font-size: 10pt; color: #64748B; padding-top: 6px; }
QLabel#cmpArrow { font-size: 22pt; color: #94A3B8; padding: 0 4px; }
QLabel#listCard { background: white; border: 1px solid #E2E8F0; border-radius: 12px; padding: 8px; }
QPushButton#sheetOption { background: transparent; border: none; border-radius: 8px; padding: 9px 12px;
                          text-align: left; font-size: 10.5pt; color: #64748B; }
QPushButton#sheetOption:hover { background: #F1F5F9; }
QPushButton#sheetOption:checked { background: #EFF6FF; color: #1E3A8A; }
QComboBox#bigSelect { background: white; border: 1px solid #CBD5E1; border-radius: 10px; padding: 10px 14px;
                      font-size: 11pt; color: #0F172A; min-height: 22px; }
QComboBox#bigSelect:hover { border-color: #2563EB; }
QComboBox#bigSelect:on { border-color: #2563EB; }
QComboBox#bigSelect::drop-down { border: none; width: 34px; }
QComboBox#bigSelect QAbstractItemView { background: white; border: 1px solid #CBD5E1; border-radius: 8px;
                                        padding: 6px; outline: none; font-size: 11pt;
                                        selection-background-color: #DBEAFE; selection-color: #0F172A; }
QPushButton#secondary { background: white; color: #2563EB; border: 2px solid #2563EB; padding: 10px 18px;
                        font-weight: 600; font-size: 11pt; }
QPushButton#secondary:hover { background: #EFF6FF; }
QPushButton#secondary:disabled { color: #94A3B8; border-color: #CBD5E1; }
QProgressBar { border: 1px solid #E2E8F0; border-radius: 6px; text-align: center; height: 20px; background: #F1F5F9; }
QProgressBar::chunk { background: #2563EB; border-radius: 6px; }
"""
