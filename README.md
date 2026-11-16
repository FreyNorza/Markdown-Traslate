# Markdown Converter & LaTeX Word Studio 📄➡️📝

<div align="center">

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![GUI](https://img.shields.io/badge/GUI-PySide6%20%2F%20Qt6-41CD52?style=for-the-badge&logo=qt&logoColor=white)](https://doc.qt.io/qtforpython/)
[![Engine](https://img.shields.io/badge/Engine-Pandoc%20%2B%20MarkItDown-800080?style=for-the-badge)](https://pandoc.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%2F%2011-0078D4?style=for-the-badge&logo=windows&logoColor=white)](https://microsoft.com)
[![License](https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge)](LICENSE)

**Aplicación de escritorio para Windows que convierte cualquier documento a Markdown y transforma Markdown con fórmulas LaTeX a documentos Word (.docx) profesionales con formato nativo OMML.**

[Descargar Ejecutable (.exe)](#-descarga-rápida-para-usuarios-sin-python) • [Características](#-funcionalidades-principales) • [Instalación para Desarrolladores](#-instalación-para-desarrolladores) • [Licencia](#-licencia)

</div>

---

## 💡 ¿Por qué nace este proyecto?

Al convertir apuntes o documentos Markdown con matemáticas (`$...$`, `$$...$$`) o diagramas a Word, la mayoría de herramientas fallan:
- Las fórmulas se convierten en **imágenes borrosas** o quedan como **texto plano roto**.
- Los títulos se tiñen de color azul por defecto de Office.
- Las imágenes vectoriales (SVG) no cargan.

**Markdown Converter** resuelve esto unificando en una interfaz gráfica moderna:
1. **Microsoft MarkItDown** para extraer y convertir más de 15 tipos de documentos a Markdown limpio.
2. **Pandoc + Postprocesador OpenXML** para convertir Markdown a Word generando **ecuaciones matemáticas nativas de Word (OMML) 100% editables**, imágenes incrustadas y tipografía académica **Times New Roman** en negrita y negro estricto.

---

## ✨ Funcionalidades Principales

### 1. 📄 Pestaña: Cualquier Archivo ➔ Markdown
Convierte prácticamente cualquier archivo a formato `.md` estándar:
- **Formatos soportados:** PDF, DOCX, PPTX, XLSX, HTML, imágenes (PNG, JPG, WebP), audio con transcripción (MP3, WAV), EPUB, ZIP, CSV, JSON, XML.
- **Vista previa integrada:** Visualiza el Markdown resultante al instante y cópialo con el botón **"📋 Copiar texto"**.
- **Apertura directa:** Accede al archivo `.md` generado o a su carpeta contenedora en un clic.

### 2. 📝 Pestaña: Markdown ➔ Word (.docx con LaTeX)
Diseñado específicamente para estudiantes, ingenieros, matemáticos y científicos:
- **Ecuaciones LaTeX Nativas (OMML):** Transforma sintaxis `$x$`, `$$\frac{a}{b}$$`, `\(...\)`, `\[...\]` y entornos `\begin{equation}` en ecuaciones reales de Word, editables con el editor de ecuaciones de Office.
- **Gráficos e Imágenes:** Soporte para PNG, JPG y vectoriales SVG.
- **Búsqueda inteligente de imágenes:** Si una imagen tiene una ruta relativa rota (ej. figuras en subcarpetas), el conversor la localiza automáticamente.
- **Formato Profesional Estricto:**
  - Tipografía **Times New Roman** en todo el documento (cuerpo en 12 pt).
  - Títulos y subtítulos en **negrita y color negro puro (`#000000`)**, eliminando la tonalidad azul por defecto de Word.
- **Carátula institucional opcional:** agrega una portada en la página 1 con universidad, facultad, escuela, título, integrantes, docente, asignatura, semestre, ciudad, año y logo. Los datos quedan guardados para reutilizarlos en próximas conversiones.
- **Tablas y Listas:** Preserva tablas complejas, listas de tareas, notas al pie y código resaltado.

### 3. 🎨 Experiencia de Usuario (UI/UX)
- **Indicador de modo en alto contraste:** Barra superior que muestra claramente la dirección de la conversión.
- **Botón `✖ Quitar archivo`:** Descarta la selección actual para elegir otro archivo sin reiniciar la app.
- **Botón `✖ Cancelar conversión`:** Permite cancelar procesos pesados en cualquier momento de manera segura.
- **Drag & Drop:** Arrastra y suelta archivos directamente sobre la aplicación.

---

## 🚀 Descarga Rápida para Usuarios (Sin Python)

Si solo deseas usar el programa en Windows sin programar ni instalar dependencias:

1. Ve a la sección de [Releases de GitHub](https://github.com/FreyNorza/Markdown-Traslate/releases).
2. Descarga el archivo ejecutable **`markdown-converter.exe`**.
3. Asegúrate de tener instalado [Pandoc](https://pandoc.org/installing.html) en tu sistema (o coloca `pandoc.exe` en la misma carpeta del ejecutable).
4. ¡Haz doble clic en `markdown-converter.exe` y listo!

---

## 🛠️ Instalación para Desarrolladores

Si deseas clonar el código, modificarlo o ejecutarlo desde la terminal:

### Requisitos previos
- Windows 10 o 11 (64-bit)
- Python 3.10 o superior
- [Pandoc](https://pandoc.org/installing.html) en el PATH del sistema

### Paso a paso

```bat
# 1. Clonar el repositorio
git clone https://github.com/FreyNorza/Markdown-Traslate.git
cd Markdown-Traslate

# 2. Crear y activar entorno virtual
python -m venv .venv
.venv\Scripts\activate

# 3. Instalar dependencias requeridas
pip install -r requirements.txt

# 4. Iniciar la aplicación
python main.py
```

---

## 📦 Compilar el Ejecutable (.exe)

El proyecto incluye un script automatizado para generar un único ejecutable portable mediante PyInstaller:

```bat
build.bat
```

El resultado se generará en la carpeta `dist/markdown-converter.exe`.

---

## 📁 Estructura del Repositorio

```
├── main.py                        # Punto de entrada de la aplicación
├── requirements.txt               # Dependencias Python
├── build.bat                      # Script para compilar el ejecutable (.exe)
├── markdown-converter.spec        # Configuración optimizada de PyInstaller
├── LICENSE                        # Licencia MIT de código abierto
├── README.md                      # Documentación del proyecto
├── ui/
│   ├── main_window.py             # Ventana principal con tabs y badges de estado
│   ├── dialogs/
│   │   └── cover_dialog.py        # Formulario modal para configurar la carátula
│   ├── panels/
│   │   ├── to_markdown_panel.py   # Panel Archivo → Markdown con vista previa
│   │   └── to_word_panel.py       # Panel Markdown → Word (Pandoc / OMML)
│   └── workers/
│       └── conversion_worker.py   # QThread con soporte de cancelación
├── converters/
│   ├── markitdown_converter.py    # Conversión con Microsoft markitdown
│   ├── cover_page.py              # Generación y persistencia de la carátula
│   └── mdtranscode_converter.py   # Conversión con Pandoc + formateo OpenXML
├── utils/
│   └── resource_path.py           # Helper de rutas para desarrollo y PyInstaller
└── assets/
    └── icon.ico                   # Icono de la aplicación
```

---

## 🤝 Contribuciones

¡Las contribuciones, reportes de bugs y sugerencias de mejoras son bienvenidas!
1. Haz un Fork del proyecto.
2. Crea una rama para tu característica (`git checkout -b feature/NuevaMejora`).
3. Haz commit de tus cambios (`git commit -m 'feat: agrega nueva mejora'`).
4. Sube tu rama (`git push origin feature/NuevaMejora`).
5. Abre un **Pull Request**.

---

## 📄 Licencia

Este proyecto está bajo la Licencia **MIT** — consulta el archivo [LICENSE](LICENSE) para más detalles.

---

<div align="center">
Hecho con ❤️ por <a href="https://github.com/FreyNorza">FreyNorza</a>
</div>
