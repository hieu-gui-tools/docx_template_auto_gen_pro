"""Entry point cho Word Template Pro."""
from __future__ import annotations

import sys
from pathlib import Path
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPalette, QColor

from word_template_pro.database import init_db
from word_template_pro.main_window import MainWindow


def main():
    init_db()
    app = QApplication(sys.argv)
    app.setApplicationName("Word Template Pro")
    app.setStyle("Fusion")

    # Dark palette base
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor("#0d1020"))
    palette.setColor(QPalette.ColorRole.WindowText, QColor("#c8d0e8"))
    palette.setColor(QPalette.ColorRole.Base, QColor("#1a1f35"))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#1e2540"))
    palette.setColor(QPalette.ColorRole.Text, QColor("#e0e8ff"))
    palette.setColor(QPalette.ColorRole.Button, QColor("#1a2040"))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor("#aabbff"))
    palette.setColor(QPalette.ColorRole.Highlight, QColor("#2d3a8a"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ddeeff"))
    app.setPalette(palette)

    icon_path = Path(__file__).parent.parent.parent / "icon.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    win = MainWindow()
    win.show()
    win.check_missing_templates()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
