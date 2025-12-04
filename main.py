"""
main.py

Entry point for the Cable Tray Calculator.

Usage:
    python main.py

Requirements:
    pip install pyqt5
"""

import sys

from PyQt5 import QtWidgets

from gui import CableTrayCalculator


def main() -> None:
    """
    Create the QApplication and show the main Cable Tray Calculator window.
    """
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName("Cable Tray Calculator")

    # Use a consistent base style; QSS in the window will do the rest.
    app.setStyle("Fusion")

    window = CableTrayCalculator()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
