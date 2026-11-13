"""
QThread genérico para ejecutar conversiones en background.
Reutilizable por ambos paneles (ToMarkdown y ToWord).
"""
from __future__ import annotations
from typing import Callable

from PySide6.QtCore import QThread, Signal


class ConversionWorker(QThread):
    """
    Ejecuta una función de conversión en un hilo secundario.

    Señales:
        finished(result): emitida cuando la conversión termina con éxito.
                          'result' es el valor de retorno de convert_fn.
        error(message):   emitida cuando la conversión lanza una excepción.
    """

    finished = Signal(object)   # objeto genérico (str o ConversionResult)
    error = Signal(str)         # mensaje de error
    cancelled = Signal()        # emitida cuando el usuario cancela la conversión

    def __init__(
        self,
        convert_fn: Callable,
        input_path: str,
        output_path: str,
        parent=None,
    ):
        super().__init__(parent)
        self._convert_fn = convert_fn
        self._input_path = input_path
        self._output_path = output_path
        self._cancelled = False

    # ------------------------------------------------------------------
    # Interfaz pública
    # ------------------------------------------------------------------

    def cancel(self) -> None:
        """Marca la conversión como cancelada. El hilo terminará pronto."""
        self._cancelled = True
        self.requestInterruption()
        self.cancelled.emit()

    # ------------------------------------------------------------------
    # Implementación QThread
    # ------------------------------------------------------------------

    def run(self) -> None:
        try:
            result = self._convert_fn(self._input_path, self._output_path)
            if not self._cancelled:
                self.finished.emit(result)
        except Exception as exc:
            if not self._cancelled:
                self.error.emit(str(exc))
