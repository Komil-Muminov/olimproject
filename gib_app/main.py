import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtGui import QIcon  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from styles import STYLE  # noqa: E402
from ui import MainWindow  # noqa: E402


def resource(name):
    """Путь к файлу из assets — и при запуске из исходников, и внутри exe."""
    base = getattr(sys, "_MEIPASS", os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    return os.path.join(base, "assets", name)


def main():
    if sys.platform == "win32":
        # своя иконка на панели задач, а не иконка Python
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Olim.GibDistributor")
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.styleHints().setColorScheme(Qt.ColorScheme.Light)   # светлая тема и при тёмной Windows
    app.setPalette(app.style().standardPalette())
    app.setStyleSheet(STYLE)
    app.setWindowIcon(QIcon(resource("icon.png")))
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
