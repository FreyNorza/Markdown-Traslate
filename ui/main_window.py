"""
Ventana principal de Markdown Converter.
Contiene un encabezado moderno con indicador de modo en alto contraste
y un QTabWidget estilizado con dos pestañas de conversión.
"""
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QFrame, QSizePolicy
)
from PySide6.QtGui import QIcon, QFont
from PySide6.QtCore import Qt

from ui.panels.to_markdown_panel import ToMarkdownPanel
from ui.panels.to_word_panel import ToWordPanel
from utils.resource_path import resource_path


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Markdown Converter")
        self.setMinimumSize(780, 600)
        self.resize(920, 680)

        # Icono de ventana
        icon_path = resource_path("assets/icon.ico")
        self.setWindowIcon(QIcon(icon_path))

        self._setup_ui()

    def _setup_ui(self) -> None:
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        root_layout = QVBoxLayout(central_widget)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(12)

        # -------------------------------------------------------------
        # 1. ENCABEZADO SUPERIOR MODERNO
        # -------------------------------------------------------------
        header_card = QFrame()
        header_card.setObjectName("HeaderCard")
        header_card.setStyleSheet("""
            #HeaderCard {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1E293B, stop:1 #0F172A);
                border-radius: 10px;
                padding: 12px;
            }
        """)
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(16, 10, 16, 10)

        # Título y subtítulo
        title_box = QVBoxLayout()
        title_box.setSpacing(2)

        app_title = QLabel("Markdown Converter")
        app_title.setStyleSheet("color: #FFFFFF; font-size: 18px; font-weight: 800; letter-spacing: 0.5px;")
        title_box.addWidget(app_title)

        app_subtitle = QLabel("Conversor bidireccional: Documentos a Markdown & Markdown a Word (LaTeX / OMML)")
        app_subtitle.setStyleSheet("color: #94A3B8; font-size: 12px;")
        title_box.addWidget(app_subtitle)

        header_layout.addLayout(title_box)
        header_layout.addStretch()

        # Insignia de modo activo (Alto contraste)
        self._mode_badge = QLabel()
        self._mode_badge.setObjectName("ModeBadge")
        header_layout.addWidget(self._mode_badge)

        root_layout.addWidget(header_card)

        # -------------------------------------------------------------
        # 2. SISTEMA DE PESTAÑAS (TABS) CON ALTO CONTRASTE
        # -------------------------------------------------------------
        self._tabs = QTabWidget()
        self._tabs.setObjectName("MainTabs")

        self._panel_md = ToMarkdownPanel()
        self._panel_word = ToWordPanel()

        self._tabs.addTab(self._panel_md, "  📄  Archivo ➔ Markdown (.md)  ")
        self._tabs.addTab(self._panel_word, "  📝  Markdown ➔ Word (.docx con LaTeX)  ")

        self._tabs.currentChanged.connect(self._on_tab_changed)

        # Estilo de pestañas de alto contraste y aspecto moderno
        self._tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 2px solid #E2E8F0;
                border-radius: 8px;
                background-color: #FFFFFF;
                top: -1px;
            }
            QTabBar::tab {
                background-color: #F1F5F9;
                color: #475569;
                padding: 10px 22px;
                margin-right: 4px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-size: 13px;
                font-weight: 600;
                border: 1px solid #CBD5E1;
                border-bottom: none;
            }
            QTabBar::tab:hover {
                background-color: #E2E8F0;
                color: #0F172A;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF;
                color: #0284C7;
                font-weight: 700;
                border: 2px solid #0284C7;
                border-bottom: 3px solid #FFFFFF;
            }
        """)

        root_layout.addWidget(self._tabs, stretch=1)

        # Inicializar estado del modo
        self._on_tab_changed(0)

    def _on_tab_changed(self, index: int) -> None:
        """Actualiza el indicador visual de alto contraste según el modo activo."""
        if index == 0:
            self._mode_badge.setText("  MODO ACTIVO: CUALQUIER ARCHIVO ➔ MARKDOWN  ")
            self._mode_badge.setStyleSheet("""
                background-color: #0284C7;
                color: #FFFFFF;
                font-size: 11px;
                font-weight: 800;
                padding: 6px 12px;
                border-radius: 12px;
                border: 1px solid #38BDF8;
            """)
        else:
            self._mode_badge.setText("  MODO ACTIVO: MARKDOWN ➔ WORD (LATEX + TIMES NEW ROMAN)  ")
            self._mode_badge.setStyleSheet("""
                background-color: #0D9488;
                color: #FFFFFF;
                font-size: 11px;
                font-weight: 800;
                padding: 6px 12px;
                border-radius: 12px;
                border: 1px solid #2DD4BF;
            """)
