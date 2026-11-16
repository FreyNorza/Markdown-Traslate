"""
Ventana modal para configurar los datos de la carátula institucional.
Los valores se guardan en disco para reutilizarlos en próximas conversiones.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QLineEdit,
    QPlainTextEdit, QPushButton, QLabel, QFileDialog, QFrame,
    QMessageBox,
)

from converters.cover_page import CoverData


_LOGO_FILTER = "Imágenes (*.png *.jpg *.jpeg *.bmp *.gif *.webp);;Todos los archivos (*.*)"


class CoverDialog(QDialog):
    """Formulario editable de la carátula."""

    def __init__(self, data: CoverData | None = None, parent=None):
        super().__init__(parent)
        self._data = data or CoverData(year=str(datetime.now().year))
        self.setWindowTitle("Configurar carátula")
        self.setFixedWidth(680)
        self.setModal(True)
        self._setup_ui()
        self._load_data(self._data)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 14, 18, 14)
        root.setSpacing(10)

        title = QLabel("Datos de la carátula")
        title.setStyleSheet("font-size: 17px; font-weight: 800; color: #0F172A;")
        root.addWidget(title)

        subtitle = QLabel(
            "Estos datos se guardarán y reutilizarán automáticamente en las próximas conversiones."
        )
        subtitle.setStyleSheet("color: #64748B; font-size: 12px;")
        subtitle.setWordWrap(True)
        root.addWidget(subtitle)

        form_card = QFrame()
        form_card.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
            }
        """)
        form = QFormLayout(form_card)
        form.setContentsMargins(16, 12, 16, 12)
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(6)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self._institution = self._add_line(
            form, "Institución:", "UNIVERSIDAD ANDINA DEL CUSCO"
        )
        self._faculty = self._add_line(
            form, "Facultad:", "FACULTAD DE INGENIERÍA Y ARQUITECTURA"
        )
        self._school = self._add_line(
            form, "Escuela / Carrera:", "ESCUELA PROFESIONAL DE INGENIERÍA DE SISTEMAS"
        )
        self._title = self._add_line(
            form, "Título del trabajo:", "Título del informe o actividad"
        )

        self._students = QPlainTextEdit()
        self._students.setPlaceholderText("Un integrante por línea")
        self._students.setFixedHeight(64)
        self._students.setStyleSheet(self._input_style())
        form.addRow("Integrantes:", self._students)

        self._teacher = self._add_line(
            form, "Docente:", "Nombre del docente"
        )
        self._subject = self._add_line(
            form, "Asignatura:", "Nombre de la asignatura"
        )
        self._semester = self._add_line(
            form, "Semestre académico:", "2026-2"
        )

        row = QHBoxLayout()
        row.setSpacing(8)
        self._city = QLineEdit()
        self._city.setPlaceholderText("CUSCO - PERÚ")
        self._city.setFixedHeight(30)
        self._city.setStyleSheet(self._input_style())
        self._year = QLineEdit()
        self._year.setPlaceholderText(str(datetime.now().year))
        self._year.setFixedWidth(110)
        self._year.setFixedHeight(30)
        self._year.setStyleSheet(self._input_style())
        row.addWidget(self._city, stretch=1)
        row.addWidget(self._year)
        form.addRow("Ciudad y año:", row)

        logo_row = QHBoxLayout()
        logo_row.setSpacing(8)
        self._logo_edit = QLineEdit()
        self._logo_edit.setReadOnly(True)
        self._logo_edit.setPlaceholderText("Sin logo seleccionado")
        self._logo_edit.setFixedHeight(30)
        self._logo_edit.setStyleSheet(self._input_style())
        logo_row.addWidget(self._logo_edit, stretch=1)

        self._btn_logo = QPushButton("Examinar...")
        self._btn_logo.setFixedHeight(32)
        self._btn_logo.setStyleSheet(self._secondary_button_style())
        self._btn_logo.clicked.connect(self._pick_logo)
        logo_row.addWidget(self._btn_logo)

        self._btn_clear_logo = QPushButton("Quitar")
        self._btn_clear_logo.setFixedHeight(32)
        self._btn_clear_logo.setStyleSheet(self._secondary_button_style())
        self._btn_clear_logo.clicked.connect(self._clear_logo)
        logo_row.addWidget(self._btn_clear_logo)
        form.addRow("Logo:", logo_row)

        self._logo_preview = QLabel("Sin logo")
        self._logo_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._logo_preview.setFixedSize(140, 56)
        self._logo_preview.setStyleSheet("""
            QLabel {
                border: 1px dashed #CBD5E1;
                border-radius: 6px;
                color: #94A3B8;
                font-size: 11px;
                background-color: #FFFFFF;
            }
        """)
        form.addRow("", self._logo_preview)

        root.addWidget(form_card)

        hint = QLabel(
            "Deja un campo vacío para omitirlo de la carátula. "
            "El título es obligatorio."
        )
        hint.setStyleSheet("color: #94A3B8; font-size: 11px;")
        hint.setWordWrap(True)
        root.addWidget(hint)

        buttons = QHBoxLayout()
        buttons.addStretch()

        self._btn_cancel = QPushButton("Cancelar")
        self._btn_cancel.setFixedHeight(38)
        self._btn_cancel.setMinimumWidth(110)
        self._btn_cancel.setStyleSheet(self._secondary_button_style())
        self._btn_cancel.clicked.connect(self.reject)
        buttons.addWidget(self._btn_cancel)

        self._btn_save = QPushButton("💾  Guardar y usar")
        self._btn_save.setFixedHeight(38)
        self._btn_save.setMinimumWidth(160)
        self._btn_save.setDefault(True)
        self._btn_save.setStyleSheet("""
            QPushButton {
                background-color: #0D9488;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                font-weight: 700;
                font-size: 13px;
                padding: 0 18px;
            }
            QPushButton:hover {
                background-color: #0F766E;
            }
            QPushButton:pressed {
                background-color: #115E59;
            }
        """)
        self._btn_save.clicked.connect(self._on_accept)
        buttons.addWidget(self._btn_save)

        root.addLayout(buttons)

    def _add_line(self, form: QFormLayout, label: str, placeholder: str) -> QLineEdit:
        line = QLineEdit()
        line.setPlaceholderText(placeholder)
        line.setFixedHeight(30)
        line.setStyleSheet(self._input_style())
        form.addRow(label, line)
        return line

    @staticmethod
    def _input_style() -> str:
        return """
            QLineEdit, QPlainTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 5px;
                padding: 6px 8px;
                font-size: 12px;
                color: #0F172A;
                selection-background-color: #99F6E4;
            }
            QLineEdit:focus, QPlainTextEdit:focus {
                border: 2px solid #0D9488;
            }
        """

    @staticmethod
    def _secondary_button_style() -> str:
        return """
            QPushButton {
                background-color: #E2E8F0;
                color: #1E293B;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                font-size: 12px;
                font-weight: 600;
                padding: 0 12px;
            }
            QPushButton:hover {
                background-color: #CBD5E1;
            }
        """

    # ------------------------------------------------------------------
    # Datos
    # ------------------------------------------------------------------

    def _load_data(self, data: CoverData) -> None:
        self._institution.setText(data.institution)
        self._faculty.setText(data.faculty)
        self._school.setText(data.school)
        self._title.setText(data.title)
        self._students.setPlainText("\n".join(data.students))
        self._teacher.setText(data.teacher)
        self._subject.setText(data.subject)
        self._semester.setText(data.semester)
        self._city.setText(data.city)
        self._year.setText(data.year or str(datetime.now().year))
        self._logo_edit.setText(data.logo_path)
        self._update_logo_preview(data.logo_path)

    def data(self) -> CoverData:
        students = [
            line.strip()
            for line in self._students.toPlainText().splitlines()
            if line.strip()
        ]
        return CoverData(
            institution=self._institution.text().strip(),
            faculty=self._faculty.text().strip(),
            school=self._school.text().strip(),
            title=self._title.text().strip(),
            students=students,
            teacher=self._teacher.text().strip(),
            subject=self._subject.text().strip(),
            semester=self._semester.text().strip(),
            city=self._city.text().strip(),
            year=self._year.text().strip(),
            logo_path=self._logo_edit.text().strip(),
        )

    # ------------------------------------------------------------------
    # Logo
    # ------------------------------------------------------------------

    def _pick_logo(self) -> None:
        initial = self._logo_edit.text().strip()
        start_dir = str(Path(initial).parent) if initial else ""
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar logo institucional", start_dir, _LOGO_FILTER
        )
        if path:
            self._logo_edit.setText(path)
            self._update_logo_preview(path)

    def _clear_logo(self) -> None:
        self._logo_edit.clear()
        self._update_logo_preview("")

    def _update_logo_preview(self, path: str) -> None:
        if not path or not Path(path).exists():
            self._logo_preview.setPixmap(QPixmap())
            self._logo_preview.setText("Sin logo")
            return
        pixmap = QPixmap(path)
        if pixmap.isNull():
            self._logo_preview.setPixmap(QPixmap())
            self._logo_preview.setText("No se pudo cargar")
            return
        self._logo_preview.setText("")
        scaled = pixmap.scaled(
            self._logo_preview.width() - 8,
            self._logo_preview.height() - 8,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self._logo_preview.setPixmap(scaled)

    # ------------------------------------------------------------------
    # Validación
    # ------------------------------------------------------------------

    def _on_accept(self) -> None:
        data = self.data()
        if not data.title:
            QMessageBox.warning(
                self,
                "Falta el título",
                "Escribe el título del trabajo o actividad para poder generar la carátula.",
            )
            self._title.setFocus()
            return

        if data.logo_path and not Path(data.logo_path).exists():
            QMessageBox.warning(
                self,
                "Logo no encontrado",
                "La imagen del logo ya no existe en la ruta indicada.\n"
                "Selecciona otra imagen o quita el logo.",
            )
            return

        self._data = data
        self.accept()

    def cover_data(self) -> CoverData:
        return self._data
