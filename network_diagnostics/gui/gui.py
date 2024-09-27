import sys

from darkdetect import theme
from PySide6.QtWidgets import QApplication, QMessageBox
from qdarktheme import load_palette, load_stylesheet

from ..utils import is_running_as_admin
from .main_window import MainWindow


def gui_main():
    app = QApplication(sys.argv)

    if not is_running_as_admin():
        QMessageBox.critical(
            None,
            'Error',
            'This application requires administrator privileges to work. Run as administrator to proceed.'
        )
        sys.exit(1)

    app_theme = 'light' if theme() == 'Light' else 'dark'
    stylesheet = load_stylesheet(app_theme)
    app.setStyleSheet(stylesheet)
    palette = load_palette(app_theme)
    app.setPalette(palette)

    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec())
