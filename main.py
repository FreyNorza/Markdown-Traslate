"""
Punto de entrada de Markdown Converter.
Instancia QApplication, aplica tema Fusion y abre la ventana principal.
"""
import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon

from ui.main_window import MainWindow
from utils.resource_path import resource_path


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setApplicationName("Markdown Converter")
    app.setApplicationVersion("1.0.0")

    icon_path = resource_path("assets/icon.ico")
    app.setWindowIcon(QIcon(icon_path))

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
