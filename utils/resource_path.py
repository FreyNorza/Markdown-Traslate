"""
Resuelve rutas de recursos tanto en modo desarrollo como dentro del .exe
empaquetado con PyInstaller (sys._MEIPASS).
"""
import sys
import os


def resource_path(relative_path: str) -> str:
    """
    Devuelve la ruta absoluta al recurso dado.

    - En modo desarrollo: relativa al directorio raíz del proyecto.
    - En modo empaquetado (.exe): relativa a sys._MEIPASS (carpeta temporal).
    """
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        # Raíz del proyecto: directorio padre de utils/
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)
