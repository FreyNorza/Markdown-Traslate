# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec para Markdown Converter.
Genera un único .exe portable (--onefile) sin ventana de consola.
Optimizado: solo los módulos de PySide6 que realmente se usan.
"""

from PyInstaller.utils.hooks import collect_all, collect_data_files

# Recolectar todo lo de markitdown (datos, binarios, hidden imports)
markitdown_datas, markitdown_binaries, markitdown_hiddenimports = collect_all('markitdown')

# Recolectar datos de magika (detector de tipos de archivo de markitdown)
magika_datas, magika_binaries, magika_hiddenimports = collect_all('magika')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=markitdown_binaries + magika_binaries,
    datas=[
        ('assets/', 'assets/'),
        *markitdown_datas,
        *magika_datas,
    ],
    hiddenimports=[
        # --- markitdown y sus dependencias dinámicas ---
        *markitdown_hiddenimports,
        *magika_hiddenimports,
        # --- PySide6: solo módulos usados ---
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        # --- Dependencias de markitdown (imports dinámicos) ---
        'pdfminer',
        'pdfplumber',
        'mammoth',
        'openpyxl',
        'pandas',
        'PIL',
        'pptx',
        'bs4',
        'lxml',
        'lxml.etree',
        # --- Post-procesado de formato (python-docx) ---
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.shared',
        'docx.image',
        'docx.image.image',
        'docx.enum.table',
        'docx.enum.text',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'scipy',
        'IPython',
        'notebook',
        'jupyter',
        # Módulos de PySide6 no usados
        'PySide6.Qt3DAnimation',
        'PySide6.Qt3DCore',
        'PySide6.Qt3DExtras',
        'PySide6.Qt3DInput',
        'PySide6.Qt3DLogic',
        'PySide6.Qt3DRender',
        'PySide6.QtAxContainer',
        'PySide6.QtBluetooth',
        'PySide6.QtCharts',
        'PySide6.QtDataVisualization',
        'PySide6.QtDesigner',
        'PySide6.QtHelp',
        'PySide6.QtLocation',
        'PySide6.QtMultimedia',
        'PySide6.QtMultimediaWidgets',
        'PySide6.QtNfc',
        'PySide6.QtPdf',
        'PySide6.QtPdfWidgets',
        'PySide6.QtPositioning',
        'PySide6.QtPrintSupport',
        'PySide6.QtQml',
        'PySide6.QtQuick',
        'PySide6.QtQuickControls2',
        'PySide6.QtQuickWidgets',
        'PySide6.QtRemoteObjects',
        'PySide6.QtScxml',
        'PySide6.QtSensors',
        'PySide6.QtSerialBus',
        'PySide6.QtSerialPort',
        'PySide6.QtSpatialAudio',
        'PySide6.QtSql',
        'PySide6.QtStateMachine',
        'PySide6.QtSvg',
        'PySide6.QtSvgWidgets',
        'PySide6.QtTest',
        'PySide6.QtUiTools',
        'PySide6.QtWebChannel',
        'PySide6.QtWebEngineCore',
        'PySide6.QtWebEngineQuick',
        'PySide6.QtWebEngineWidgets',
        'PySide6.QtWebSockets',
        'PySide6.QtXml',
        'PySide6.QtAsyncio',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='markdown-converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # Sin ventana de consola
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.ico',
)
