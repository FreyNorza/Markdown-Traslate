"""
Panel de conversión Archivo → Markdown (Unit 02).
Permite seleccionar cualquier archivo soportado via botón o drag & drop,
convertirlo a .md con Microsoft markitdown, ver vista previa y copiar o abrir el resultado.
Incluye controles explícitos de cancelar selección y cancelar conversión.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QPlainTextEdit, QProgressBar, QFrame, QSizePolicy,
    QApplication
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont

from ui.workers.conversion_worker import ConversionWorker
from converters.markitdown_converter import convert_to_markdown


def _format_size(num_bytes: int) -> str:
    """Formatea bytes a KB o MB legibles."""
    if num_bytes < 1024:
        return f"{num_bytes} B"
    elif num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    else:
        return f"{num_bytes / (1024 * 1024):.2f} MB"


class DropZone(QLabel):
    """Zona de arrastre estilizada que acepta archivos y emite file_dropped(path)."""

    file_dropped = Signal(str)

    def __init__(self, accepted_extensions: list[str] | None = None, parent=None):
        super().__init__(parent)
        self._accepted_extensions = accepted_extensions
        self._setup_ui()

    def _setup_ui(self) -> None:
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumHeight(130)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._set_idle_style()

    def _set_idle_style(self) -> None:
        self.setText("📥  Arrastra y suelta aquí cualquier archivo (PDF, DOCX, XLSX, PPTX, etc.)\no usa el botón de selección arriba")
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #94A3B8;
                border-radius: 10px;
                color: #475569;
                font-size: 14px;
                font-weight: 600;
                background-color: #F8FAFC;
                padding: 24px;
            }
            QLabel:hover {
                border-color: #0284C7;
                background-color: #F0F9FF;
                color: #0369A1;
            }
        """)

    def _set_hover_style(self) -> None:
        self.setStyleSheet("""
            QLabel {
                border: 2px solid #0284C7;
                border-radius: 10px;
                color: #0369A1;
                font-size: 14px;
                font-weight: 700;
                background-color: #E0F2FE;
                padding: 24px;
            }
        """)

    def _set_reject_style(self) -> None:
        self.setStyleSheet("""
            QLabel {
                border: 2px solid #EF4444;
                border-radius: 10px;
                color: #B91C1C;
                font-size: 13px;
                font-weight: 600;
                background-color: #FEF2F2;
                padding: 24px;
            }
        """)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                path = urls[0].toLocalFile()
                ext = Path(path).suffix.lower()
                if self._accepted_extensions is None or ext in self._accepted_extensions:
                    event.acceptProposedAction()
                    self._set_hover_style()
                    return
                else:
                    self.setText(f"❌ Formato no aceptado: {ext}")
                    self._set_reject_style()
        event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self._set_idle_style()

    def dropEvent(self, event: QDropEvent) -> None:
        self._set_idle_style()
        urls = event.mimeData().urls()
        if not urls:
            return
        path = urls[0].toLocalFile()
        ext = Path(path).suffix.lower()
        if self._accepted_extensions and ext not in self._accepted_extensions:
            self.setText(f"❌ Solo se aceptan archivos {', '.join(self._accepted_extensions)}")
            return
        event.acceptProposedAction()
        self.file_dropped.emit(path)


# Filtro de formatos para el diálogo de archivo
_FILE_FILTER = (
    "Archivos soportados ("
    "*.pdf *.docx *.doc *.pptx *.ppt *.xlsx *.xls "
    "*.html *.htm *.png *.jpg *.jpeg *.gif *.webp "
    "*.mp3 *.wav *.m4a *.zip *.epub *.csv *.json *.xml"
    ");;"
    "Todos los archivos (*.*)"
)

_PREVIEW_LIMIT = 500 * 1024  # 500 KB


class ToMarkdownPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._input_path: str = ""
        self._output_dir: str = ""
        self._worker: ConversionWorker | None = None
        self._last_output_path: str = ""
        self._setup_ui()

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(14)

        # -------------------------------------------------------------
        # 1. BARRA SUPERIOR DE HERRAMIENTAS
        # -------------------------------------------------------------
        top_card = QFrame()
        top_card.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                padding: 8px 12px;
            }
        """)
        top_layout = QHBoxLayout(top_card)
        top_layout.setContentsMargins(8, 4, 8, 4)

        self._btn_open = QPushButton("📂  Seleccionar archivo")
        self._btn_open.setFixedHeight(38)
        self._btn_open.setStyleSheet("""
            QPushButton {
                background-color: #0369A1;
                color: #FFFFFF;
                font-weight: 700;
                font-size: 13px;
                border-radius: 6px;
                padding: 0 16px;
                border: none;
            }
            QPushButton:hover {
                background-color: #0284C7;
            }
            QPushButton:pressed {
                background-color: #075985;
            }
        """)
        self._btn_open.clicked.connect(self._on_open_dialog)
        top_layout.addWidget(self._btn_open)

        top_layout.addStretch()

        dest_icon = QLabel("📁  Guardar en:")
        dest_icon.setStyleSheet("color: #334155; font-weight: 600; font-size: 12px;")
        top_layout.addWidget(dest_icon)

        self._lbl_outdir = QLabel("Misma carpeta que el original")
        self._lbl_outdir.setStyleSheet("color: #64748B; font-size: 12px; font-style: italic;")
        top_layout.addWidget(self._lbl_outdir)

        self._btn_change_dir = QPushButton("Cambiar...")
        self._btn_change_dir.setFixedHeight(30)
        self._btn_change_dir.setStyleSheet("""
            QPushButton {
                background-color: #E2E8F0;
                color: #1E293B;
                font-size: 12px;
                font-weight: 600;
                border-radius: 4px;
                padding: 0 12px;
                border: 1px solid #CBD5E1;
            }
            QPushButton:hover {
                background-color: #CBD5E1;
            }
        """)
        self._btn_change_dir.clicked.connect(self._on_change_output_dir)
        top_layout.addWidget(self._btn_change_dir)

        root.addWidget(top_card)

        # -------------------------------------------------------------
        # 2. ZONA DE DROP
        # -------------------------------------------------------------
        self._drop_zone = DropZone(accepted_extensions=None)
        self._drop_zone.file_dropped.connect(self._on_file_selected)
        root.addWidget(self._drop_zone)

        # -------------------------------------------------------------
        # 3. TARJETA DE ARCHIVO SELECCIONADO (CON BOTÓN DESCARTE)
        # -------------------------------------------------------------
        self._file_card = QFrame()
        self._file_card.setVisible(False)
        self._file_card.setStyleSheet("""
            QFrame {
                background-color: #F0F9FF;
                border: 1px solid #BAE6FD;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        file_card_layout = QHBoxLayout(self._file_card)
        file_card_layout.setContentsMargins(12, 8, 12, 8)

        file_icon = QLabel("📄")
        file_icon.setStyleSheet("font-size: 24px;")
        file_card_layout.addWidget(file_icon)

        file_info_layout = QVBoxLayout()
        file_info_layout.setSpacing(2)
        self._lbl_filename = QLabel()
        self._lbl_filename.setStyleSheet("color: #0369A1; font-size: 14px; font-weight: 700;")
        file_info_layout.addWidget(self._lbl_filename)

        self._lbl_filepath = QLabel()
        self._lbl_filepath.setStyleSheet("color: #64748B; font-size: 11px;")
        file_info_layout.addWidget(self._lbl_filepath)

        file_card_layout.addLayout(file_info_layout, stretch=1)

        # Botón para descartar el archivo seleccionado
        self._btn_remove_file = QPushButton("✖  Quitar archivo")
        self._btn_remove_file.setFixedHeight(32)
        self._btn_remove_file.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                color: #EF4444;
                border: 1px solid #FCA5A5;
                border-radius: 5px;
                font-weight: 600;
                font-size: 12px;
                padding: 0 12px;
            }
            QPushButton:hover {
                background-color: #FEE2E2;
                border-color: #EF4444;
            }
        """)
        self._btn_remove_file.clicked.connect(self._on_remove_file)
        file_card_layout.addWidget(self._btn_remove_file)

        root.addWidget(self._file_card)

        # -------------------------------------------------------------
        # 4. BOTÓN DE CONVERTIR Y BOTÓN DE CANCELAR
        # -------------------------------------------------------------
        actions_layout = QHBoxLayout()
        actions_layout.setSpacing(10)

        self._btn_convert = QPushButton("⚡  Convertir a Markdown (.md)")
        self._btn_convert.setFixedHeight(44)
        self._btn_convert.setEnabled(False)
        self._btn_convert.setStyleSheet("""
            QPushButton:enabled {
                background-color: #0284C7;
                color: #FFFFFF;
                border-radius: 6px;
                font-weight: 700;
                font-size: 14px;
                border: none;
                padding: 0 24px;
            }
            QPushButton:enabled:hover {
                background-color: #0369A1;
            }
            QPushButton:disabled {
                background-color: #E2E8F0;
                color: #94A3B8;
                border-radius: 6px;
                font-size: 14px;
                border: none;
            }
        """)
        self._btn_convert.clicked.connect(self._on_convert)
        actions_layout.addWidget(self._btn_convert, stretch=2)

        # Botón de cancelar conversión en curso
        self._btn_cancel = QPushButton("✖  Cancelar conversión")
        self._btn_cancel.setFixedHeight(44)
        self._btn_cancel.setVisible(False)
        self._btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #DC2626;
                color: #FFFFFF;
                border-radius: 6px;
                font-weight: 700;
                font-size: 13px;
                border: none;
                padding: 0 18px;
            }
            QPushButton:hover {
                background-color: #B91C1C;
            }
        """)
        self._btn_cancel.clicked.connect(self._on_cancel)
        actions_layout.addWidget(self._btn_cancel, stretch=1)

        root.addLayout(actions_layout)

        # Estado y progreso
        self._status_container = QWidget()
        self._status_container.setVisible(False)
        status_vbox = QVBoxLayout(self._status_container)
        status_vbox.setContentsMargins(0, 0, 0, 0)
        status_vbox.setSpacing(4)

        self._lbl_status = QLabel("⏳  Extrayendo texto y formateando a Markdown...")
        self._lbl_status.setStyleSheet("color: #0369A1; font-size: 12px; font-weight: 600;")
        status_vbox.addWidget(self._lbl_status)

        self._progress = QProgressBar()
        self._progress.setRange(0, 0)
        self._progress.setFixedHeight(8)
        self._progress.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 4px;
                background-color: #E2E8F0;
            }
            QProgressBar::chunk {
                background-color: #0284C7;
                border-radius: 4px;
            }
        """)
        status_vbox.addWidget(self._progress)
        root.addWidget(self._status_container)

        # Mensaje de error
        self._lbl_error = QLabel()
        self._lbl_error.setStyleSheet("""
            color: #DC2626;
            background-color: #FEF2F2;
            border: 1px solid #FCA5A5;
            border-radius: 6px;
            padding: 8px 12px;
            font-size: 12px;
            font-weight: 600;
        """)
        self._lbl_error.setWordWrap(True)
        self._lbl_error.setVisible(False)
        root.addWidget(self._lbl_error)

        # -------------------------------------------------------------
        # 5. VISTA PREVIA Y ACCIONES DE RESULTADO
        # -------------------------------------------------------------
        self._preview_container = QWidget()
        self._preview_container.setVisible(False)
        preview_vbox = QVBoxLayout(self._preview_container)
        preview_vbox.setContentsMargins(0, 0, 0, 0)
        preview_vbox.setSpacing(8)

        # Fila superior de vista previa
        prev_header = QHBoxLayout()
        self._lbl_preview_title = QLabel("👁️  Vista previa del Markdown generado:")
        self._lbl_preview_title.setStyleSheet("font-weight: 700; color: #1E293B; font-size: 13px;")
        prev_header.addWidget(self._lbl_preview_title)

        prev_header.addStretch()

        self._btn_copy_text = QPushButton("📋  Copiar texto")
        self._btn_copy_text.setFixedHeight(28)
        self._btn_copy_text.setStyleSheet("""
            QPushButton {
                background-color: #E2E8F0;
                color: #334155;
                font-weight: 600;
                font-size: 12px;
                border-radius: 4px;
                padding: 0 10px;
                border: 1px solid #CBD5E1;
            }
            QPushButton:hover {
                background-color: #CBD5E1;
            }
        """)
        self._btn_copy_text.clicked.connect(self._on_copy_text)
        prev_header.addWidget(self._btn_copy_text)

        self._btn_open_result = QPushButton("📄  Abrir archivo .md")
        self._btn_open_result.setFixedHeight(28)
        self._btn_open_result.setStyleSheet("""
            QPushButton {
                background-color: #0284C7;
                color: #FFFFFF;
                font-weight: 600;
                font-size: 12px;
                border-radius: 4px;
                padding: 0 12px;
                border: none;
            }
            QPushButton:hover {
                background-color: #0369A1;
            }
        """)
        self._btn_open_result.clicked.connect(self._on_open_result)
        prev_header.addWidget(self._btn_open_result)

        self._btn_open_folder = QPushButton("📁  Abrir carpeta")
        self._btn_open_folder.setFixedHeight(28)
        self._btn_open_folder.setStyleSheet("""
            QPushButton {
                background-color: #F1F5F9;
                color: #334155;
                font-weight: 600;
                font-size: 12px;
                border-radius: 4px;
                padding: 0 10px;
                border: 1px solid #CBD5E1;
            }
            QPushButton:hover {
                background-color: #E2E8F0;
            }
        """)
        self._btn_open_folder.clicked.connect(self._on_open_folder)
        prev_header.addWidget(self._btn_open_folder)

        preview_vbox.addLayout(prev_header)

        # Editor de solo lectura
        self._preview = QPlainTextEdit()
        self._preview.setReadOnly(True)
        mono = QFont("Consolas, Courier New", 10)
        self._preview.setFont(mono)
        self._preview.setStyleSheet("""
            QPlainTextEdit {
                background-color: #F8FAFC;
                color: #0F172A;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 8px;
            }
        """)
        preview_vbox.addWidget(self._preview, stretch=1)

        root.addWidget(self._preview_container, stretch=1)

    # ------------------------------------------------------------------
    # Handlers
    # ------------------------------------------------------------------

    def _on_open_dialog(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo a convertir", "", _FILE_FILTER
        )
        if path:
            self._on_file_selected(path)

    def _on_file_selected(self, path: str) -> None:
        self._input_path = path
        p = Path(path)
        size_str = _format_size(p.stat().st_size) if p.exists() else ""

        self._lbl_filename.setText(f"{p.name}  ({size_str})")
        self._lbl_filepath.setText(str(p.resolve()))
        self._file_card.setVisible(True)
        self._drop_zone.setVisible(False)

        self._btn_convert.setEnabled(True)
        self._lbl_error.setVisible(False)
        self._preview_container.setVisible(False)

    def _on_remove_file(self) -> None:
        """Descarta el archivo seleccionado y regresa al estado inicial."""
        self._input_path = ""
        self._file_card.setVisible(False)
        self._drop_zone.setVisible(True)
        self._btn_convert.setEnabled(False)
        self._preview_container.setVisible(False)
        self._lbl_error.setVisible(False)
        self._status_container.setVisible(False)

    def _on_change_output_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self, "Seleccionar carpeta de salida", self._output_dir or ""
        )
        if folder:
            self._output_dir = folder
            self._lbl_outdir.setText(folder)

    def _on_convert(self) -> None:
        if not self._input_path:
            return

        input_path = Path(self._input_path)
        out_dir = Path(self._output_dir) if self._output_dir else input_path.parent
        output_path = str(out_dir / (input_path.stem + ".md"))
        self._last_output_path = output_path

        self._set_converting_state(True)
        self._lbl_error.setVisible(False)
        self._preview_container.setVisible(False)

        self._worker = ConversionWorker(convert_to_markdown, str(input_path), output_path)
        self._worker.finished.connect(self._on_success)
        self._worker.error.connect(self._on_error)
        self._worker.cancelled.connect(self._on_cancelled_by_user)
        self._worker.finished.connect(lambda _: self._set_converting_state(False))
        self._worker.error.connect(lambda _: self._set_converting_state(False))
        self._worker.cancelled.connect(lambda: self._set_converting_state(False))
        self._worker.start()

    def _on_cancel(self) -> None:
        """Cancela la conversión en progreso."""
        if self._worker and self._worker.isRunning():
            self._lbl_status.setText("⏹  Cancelando conversión...")
            self._worker.cancel()
            self._worker.wait(2000)
        self._set_converting_state(False)

    def _on_cancelled_by_user(self) -> None:
        self._lbl_error.setText("ℹ️  Conversión cancelada por el usuario.")
        self._lbl_error.setStyleSheet("""
            color: #475569;
            background-color: #F1F5F9;
            border: 1px solid #CBD5E1;
            border-radius: 6px;
            padding: 8px 12px;
            font-size: 12px;
            font-weight: 600;
        """)
        self._lbl_error.setVisible(True)

    def _on_success(self, content: str) -> None:
        preview_text = content
        truncated = False
        encoded = content.encode("utf-8")
        if len(encoded) > _PREVIEW_LIMIT:
            preview_text = encoded[:_PREVIEW_LIMIT].decode("utf-8", errors="ignore")
            truncated = True

        self._preview.setPlainText(preview_text)
        if truncated:
            self._preview.appendPlainText(
                "\n\n[... vista previa truncada a 500 KB. El archivo completo fue guardado en el disco.]"
            )

        self._preview_container.setVisible(True)

    def _on_error(self, message: str) -> None:
        self._lbl_error.setText(f"❌  <b>Error durante la conversión:</b><br>{message}")
        self._lbl_error.setStyleSheet("""
            color: #DC2626;
            background-color: #FEF2F2;
            border: 1px solid #FCA5A5;
            border-radius: 6px;
            padding: 8px 12px;
            font-size: 12px;
            font-weight: 600;
        """)
        self._lbl_error.setVisible(True)

    def _on_copy_text(self) -> None:
        clipboard = QApplication.clipboard()
        clipboard.setText(self._preview.toPlainText())
        self._btn_copy_text.setText("✅  ¡Copiado!")
        from PySide6.QtCore import QTimer
        QTimer.singleShot(2000, lambda: self._btn_copy_text.setText("📋  Copiar texto"))

    def _on_open_result(self) -> None:
        path = getattr(self, "_last_output_path", "")
        if path and Path(path).exists():
            if sys.platform == "win32":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        else:
            self._lbl_error.setText("❌ El archivo generado no se encontró.")
            self._lbl_error.setVisible(True)

    def _on_open_folder(self) -> None:
        if self._last_output_path:
            p = Path(self._last_output_path)
            folder = p.parent if p.exists() else Path(self._output_dir or ".")
            if sys.platform == "win32":
                if p.exists():
                    subprocess.Popen(f'explorer /select,"{p}"')
                else:
                    os.startfile(str(folder))
            else:
                subprocess.Popen(["xdg-open", str(folder)])

    def _set_converting_state(self, converting: bool) -> None:
        if converting:
            self._btn_convert.setVisible(False)
            self._btn_cancel.setVisible(True)
            self._btn_remove_file.setEnabled(False)
            self._btn_open.setEnabled(False)
            self._btn_change_dir.setEnabled(False)
            self._status_container.setVisible(True)
            self._lbl_status.setText("⏳  Extrayendo texto y formateando a Markdown...")
        else:
            self._btn_convert.setVisible(True)
            self._btn_convert.setEnabled(bool(self._input_path))
            self._btn_cancel.setVisible(False)
            self._btn_remove_file.setEnabled(True)
            self._btn_open.setEnabled(True)
            self._btn_change_dir.setEnabled(True)
            self._status_container.setVisible(False)
