"""Entry point for the Pharmacy Invoice Generator desktop app."""

from PyQt5.QtWidgets import QApplication
import sys

# Prefer the consolidated UI in app/ui_main.py
from app.ui_main import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
