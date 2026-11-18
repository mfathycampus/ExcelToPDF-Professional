from __future__ import annotations

import sys

from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from .. import tempstore
from .main_window import MainWindow


def main() -> int:
    tempstore.cleanup_stale()           # leftovers from crashed sessions
    app = QApplication(sys.argv)
    app.setApplicationName("Excel to PDF Professional")
    app.setStyle("Fusion")
    icon = Path(__file__).resolve().parents[1] / "resources" / "app.png"
    if icon.is_file():
        app.setWindowIcon(QIcon(str(icon)))
    win = MainWindow()
    win.show()
    if len(sys.argv) > 1:
        win.open_file(sys.argv[1])
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
