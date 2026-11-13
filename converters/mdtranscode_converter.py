r"""
Conversor Markdown → Word (.docx) usando Pandoc como motor.

Flujo de conversión:
  1. Pandoc convierte MD → DOCX (LaTeX math → OMML, imágenes, tablas, etc.)
  2. python-docx post-procesa el .docx para aplicar formato personalizado:
     - Fuente Times New Roman en todo el documento
     - Títulos y subtítulos en negrita y color negro puro
     - Eliminación de temas azules de Word

Pandoc maneja de forma nativa:
  - LaTeX math en todas sus notaciones:
      · Inline:  $...$, \(...\)
      · Display: $$...$$, \[...\], \begin{equation}...
  - Las ecuaciones se convierten a OMML (Office Math Markup Language),
    lo que las hace **editables directamente en Word**.
  - Imágenes: PNG, JPG, SVG (incrustadas en el .docx)
  - Tablas, encabezados, listas, código, citas, footnotes, etc.

Estrategia de búsqueda de imágenes:
  El conversor resuelve rutas de imágenes rotas buscando el archivo
  por nombre en la carpeta del .md y en sus subcarpetas (un nivel).
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import NamedTuple


# ---------------------------------------------------------------------------
# Tipos auxiliares
# ---------------------------------------------------------------------------

class ConversionResult(NamedTuple):
    output_path: str
    warnings: list[str]


# ---------------------------------------------------------------------------
# Localización de pandoc
# ---------------------------------------------------------------------------

def _find_pandoc() -> str:
    """
    Busca pandoc en:
    1. PATH del sistema
    2. Ubicación típica de instalación en Windows
    3. Junto al ejecutable de la app (bundled)

    Returns:
        Ruta al ejecutable de pandoc.

    Raises:
        FileNotFoundError si no se encuentra.
    """
    # 1. PATH
    pandoc = shutil.which("pandoc")
    if pandoc:
        return pandoc

    # 2. Ubicaciones típicas de Windows
    candidates = [
        Path(os.environ.get("LOCALAPPDATA", "")) / "Pandoc" / "pandoc.exe",
        Path(os.environ.get("PROGRAMFILES", "")) / "Pandoc" / "pandoc.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Pandoc" / "pandoc.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)

    # 3. Junto al .exe (bundled en distribución portable)
    from utils.resource_path import resource_path
    bundled = Path(resource_path("pandoc.exe"))
    if bundled.exists():
        return str(bundled)

    raise FileNotFoundError(
        "No se encontró Pandoc.\n\n"
        "Pandoc es necesario para convertir Markdown a Word.\n"
        "Descárgalo e instálalo desde:\n"
        "https://pandoc.org/installing.html\n\n"
        "También puedes colocar pandoc.exe en la misma carpeta que la aplicación."
    )


# ---------------------------------------------------------------------------
# Resolución de rutas de imágenes
# ---------------------------------------------------------------------------

def _resolve_image_paths(md_text: str, source_dir: Path) -> tuple[str, list[str]]:
    """
    Recorre las referencias de imagen del Markdown y verifica que los
    archivos existan. Si una ruta no existe, busca el archivo por nombre
    en source_dir y sus subdirectorios inmediatos.

    Returns:
        (md_text_corregido, lista_de_avisos)
    """
    warnings: list[str] = []

    # Patrón: ![alt](ruta) o ![alt](ruta "title")
    img_pattern = re.compile(r'!\[([^\]]*)\]\(([^)"\s]+)([^)]*)\)')

    def _resolve(match: re.Match) -> str:
        alt   = match.group(1)
        path  = match.group(2)
        rest  = match.group(3)   # título opcional

        # Rutas absolutas o URLs: dejar igual
        if path.startswith(("http://", "https://", "data:", "/")):
            return match.group(0)

        img_path = (source_dir / path).resolve()

        if img_path.exists():
            return match.group(0)   # ya está bien

        # Buscar el archivo por nombre en source_dir y un nivel de subcarpetas
        name = Path(path).name
        found: Path | None = None

        search_dirs = [source_dir]
        try:
            search_dirs += [p for p in source_dir.iterdir() if p.is_dir()]
        except OSError:
            pass

        for candidate in search_dirs:
            p = candidate / name
            if p.exists():
                found = p
                break

        if found:
            # Reescribir con ruta absoluta (pandoc la acepta)
            new_path = found.as_posix()
            return f"![{alt}]({new_path}{rest})"
        else:
            warnings.append(f"Imagen no encontrada: {path}")
            return match.group(0)   # pandoc mostrará su propio warning

    corrected = img_pattern.sub(_resolve, md_text)
    return corrected, warnings


# ---------------------------------------------------------------------------
# Post-procesado de formato del .docx
# ---------------------------------------------------------------------------

_FONT_NAME = "Times New Roman"

# Tamaños de fuente para cada nivel de encabezado (en puntos)
_HEADING_SIZES = {
    1: 18,   # Título principal
    2: 16,   # Subtítulo
    3: 14,   # Sub-subtítulo
    4: 13,
    5: 12,
    6: 12,
}

_BODY_SIZE = 12  # Tamaño del texto normal


def _force_xml_formatting(
    element,
    font_name: str = _FONT_NAME,
    bold: bool | None = None,
    color_black: bool = False,
    size_pt: float | None = None,
) -> None:
    """
    Modifica directamente el XML de un run (<w:r>) o estilo (<w:style>)
    para garantizar que Word no aplique temas (w:themeColor), colores azules
    por defecto de Office, o fuentes del sistema (como Calibri).
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    # En un run o estilo, las propiedades de fuente residen en <w:rPr>
    if element.tag.endswith("rPr"):
        rPr = element
    else:
        rPr = element.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            element.append(rPr)

    # 1. Fuente Times New Roman en todas las codificaciones (ascii, hAnsi, cs, eastAsia)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)

    # 2. Negrita forzada
    if bold is True:
        b = rPr.find(qn("w:b"))
        if b is None:
            b = OxmlElement("w:b")
            rPr.append(b)
        b.set(qn("w:val"), "1")

        bCs = rPr.find(qn("w:bCs"))
        if bCs is None:
            bCs = OxmlElement("w:bCs")
            rPr.append(bCs)
        bCs.set(qn("w:val"), "1")

    # 3. Color negro puro: elimina w:themeColor / w:themeShade que Word usa para azul
    if color_black:
        color = rPr.find(qn("w:color"))
        if color is None:
            color = OxmlElement("w:color")
            rPr.append(color)
        color.set(qn("w:val"), "000000")

        # ELIMINAR atributos de tema (accent1, etc.) para que Word NO pinte de azul
        for attr in ["w:themeColor", "w:themeTint", "w:themeShade"]:
            key = qn(attr)
            if key in color.attrib:
                del color.attrib[key]

    # 4. Tamaño de fuente si se especifica
    if size_pt is not None:
        sz_val = str(int(size_pt * 2))  # Word usa medios puntos (half-points)
        sz = rPr.find(qn("w:sz"))
        if sz is None:
            sz = OxmlElement("w:sz")
            rPr.append(sz)
        sz.set(qn("w:val"), sz_val)

        szCs = rPr.find(qn("w:szCs"))
        if szCs is None:
            szCs = OxmlElement("w:szCs")
            rPr.append(szCs)
        szCs.set(qn("w:val"), sz_val)


def _is_heading_or_title(para) -> bool:
    """Detecta si un párrafo es un encabezado, título o subtítulo."""
    style_name = (getattr(para.style, "name", "") or "").lower()
    style_id = (getattr(para.style, "style_id", "") or "").lower()
    return (
        style_name.startswith("heading")
        or style_id.startswith("heading")
        or style_name in ("title", "subtitle")
        or style_id in ("title", "subtitle")
        or "encabezado" in style_name
        or "título" in style_name
        or "titulo" in style_name
        or "subtítulo" in style_name
        or "subtitulo" in style_name
    )


def _get_heading_level(para) -> int | None:
    """Extrae el nivel numérico de un encabezado (1 a 6) si existe."""
    style_name = getattr(para.style, "name", "") or ""
    style_id = getattr(para.style, "style_id", "") or ""
    for text in (style_name, style_id):
        m = re.search(r"heading\s*(\d)", text, re.IGNORECASE)
        if m:
            return int(m.group(1))
        m_es = re.search(r"(?:título|titulo|encabezado)\s*(\d)", text, re.IGNORECASE)
        if m_es:
            return int(m_es.group(1))
    return None


def _postprocess_docx(docx_path: Path) -> None:
    """
    Aplica formato personalizado al .docx generado por Pandoc:
      - Fuente Times New Roman en todo el documento.
      - Títulos y subtítulos (que Word por defecto genera en azul)
        forzados a COLOR NEGRO (#000000) y EN NEGRITA.
      - Eliminación total de atributos de tema (accent1/themeColor) para evitar azul residual.
      - Tamaños jerárquicos para encabezados y texto base a 12pt.
    """
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document(str(docx_path))

    # ----- 1. Configurar y limpiar estilos del documento -----

    # Estilo Normal (base para el texto general)
    try:
        normal = doc.styles["Normal"]
        normal.font.name = _FONT_NAME
        normal.font.size = Pt(_BODY_SIZE)
        normal.font.color.rgb = RGBColor(0, 0, 0)
        _force_xml_formatting(normal.element, font_name=_FONT_NAME, bold=False, color_black=True, size_pt=_BODY_SIZE)
    except KeyError:
        pass

    # Limpiar todos los estilos de encabezado registrados en doc.styles
    for style in doc.styles:
        s_name = (style.name or "").lower()
        s_id = (getattr(style, "style_id", "") or "").lower()

        is_h = (
            s_name.startswith("heading")
            or s_id.startswith("heading")
            or s_name in ("title", "subtitle")
            or s_id in ("title", "subtitle")
            or "título" in s_name
            or "titulo" in s_name
            or "encabezado" in s_name
        )

        if is_h:
            # Determinar tamaño según nivel si es posible
            level = None
            m = re.search(r"(\d)", s_name + s_id)
            if m:
                level = int(m.group(1))

            size = _HEADING_SIZES.get(level, 14 if "subtitle" in s_name else 18)
            _force_xml_formatting(
                style.element,
                font_name=_FONT_NAME,
                bold=True,
                color_black=True,
                size_pt=size,
            )
        else:
            # Estilos regulares (párrafos, listas, tablas, etc.)
            _force_xml_formatting(style.element, font_name=_FONT_NAME, bold=None, color_black=False)

    # ----- 2. Formato directo párrafo por párrafo y run por run -----
    # Word a menudo sobreescribe los estilos con propiedades en los runs generados por Pandoc.
    # Recorremos cada párrafo para garantizar 100% que los títulos sean negros y en negrita.
    for para in doc.paragraphs:
        if _is_heading_or_title(para):
            level = _get_heading_level(para)
            heading_size = _HEADING_SIZES.get(level, 16)

            # Si el párrafo tiene pPr, asegurarse de que su rPr de párrafo también tenga negro/negrita
            pPr = para._p.get_or_add_pPr()
            rPr = pPr.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                pPr.append(rPr)
            _force_xml_formatting(rPr, font_name=_FONT_NAME, bold=True, color_black=True, size_pt=heading_size)

            # Aplicar a cada run dentro del título
            for run in para.runs:
                _force_xml_formatting(run._r, font_name=_FONT_NAME, bold=True, color_black=True, size_pt=heading_size)
        else:
            # Párrafos normales
            for run in para.runs:
                _force_xml_formatting(run._r, font_name=_FONT_NAME, bold=None, color_black=False)

    # ----- 3. Tablas -----
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    is_h = _is_heading_or_title(para)
                    for run in para.runs:
                        if is_h:
                            _force_xml_formatting(run._r, font_name=_FONT_NAME, bold=True, color_black=True)
                        else:
                            _force_xml_formatting(run._r, font_name=_FONT_NAME)

    # ----- 4. Encabezados y pies de página de sección -----
    for section in doc.sections:
        for header_footer in [section.header, section.footer]:
            if header_footer is not None:
                for para in header_footer.paragraphs:
                    for run in para.runs:
                        _force_xml_formatting(run._r, font_name=_FONT_NAME)

    doc.save(str(docx_path))


# ---------------------------------------------------------------------------
# Conversor principal
# ---------------------------------------------------------------------------

# Extensiones de Pandoc Markdown para máximo soporte de LaTeX y formato
_PANDOC_EXTENSIONS = "+".join([
    "markdown",
    # --- LaTeX math ---
    "tex_math_dollars",            # $...$ y $$...$$
    "tex_math_single_backslash",   # \(...\) y \[...\]
    "tex_math_double_backslash",   # \\(...\\) y \\[...\\]
    "raw_tex",                     # comandos LaTeX crudos (\textbf, \newpage, etc.)
    # --- Tablas ---
    "pipe_tables",                 # tablas con | separadores
    "grid_tables",                 # tablas con bordes ASCII
    "multiline_tables",            # tablas multilínea
    # --- Otros ---
    "yaml_metadata_block",         # front matter YAML (título, autor, fecha)
    "fenced_code_blocks",          # bloques de código con ```
    "footnotes",                   # notas al pie [^1]
    "strikeout",                   # ~~tachado~~
    "superscript",                 # ^superíndice^
    "subscript",                   # ~subíndice~
    "task_lists",                  # - [x] checkboxes
    "definition_lists",            # listas de definición
])


def convert_to_word(input_path: str, output_path: str) -> ConversionResult:
    """
    Convierte el archivo .md en input_path a .docx usando Pandoc,
    y luego aplica formato personalizado (Times New Roman, títulos en negrita).

    Las ecuaciones LaTeX ($...$, $$...$$, \\(...\\), \\[...\\],
    \\begin{equation}, etc.) se convierten automáticamente a
    ecuaciones OMML nativas de Word, editables con el editor
    de ecuaciones integrado.

    Args:
        input_path:  Ruta al archivo .md de entrada.
        output_path: Ruta donde se guardará el .docx.

    Returns:
        ConversionResult con la ruta de salida y lista de avisos.

    Raises:
        FileNotFoundError: Si pandoc no está disponible.
        RuntimeError:      Si pandoc termina con error.
    """
    warnings: list[str] = []
    input_path  = Path(input_path).resolve()
    output_path = Path(output_path).resolve()
    source_dir  = input_path.parent

    # --- Localizar pandoc ---
    pandoc_exe = _find_pandoc()

    # --- Leer fuente y resolver rutas de imágenes ---
    md_text = input_path.read_text(encoding="utf-8")
    md_text, img_warnings = _resolve_image_paths(md_text, source_dir)
    warnings.extend(img_warnings)

    # --- Escribir .md corregido en archivo temporal ---
    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", suffix=".md", delete=False
    ) as tmp:
        tmp.write(md_text)
        tmp_path = tmp.name

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)

        cmd = [
            pandoc_exe,
            tmp_path,
            "-o", str(output_path),
            # Formato de entrada con todas las extensiones LaTeX
            "--from", _PANDOC_EXTENSIONS,
            # Rutas de recursos: carpeta del .md original
            "--resource-path", str(source_dir),
            "--standalone",
            "--wrap=none",
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=120,
        )

        # Capturar warnings de pandoc (van a stderr)
        if result.stderr:
            for line in result.stderr.splitlines():
                line = line.strip()
                if line.startswith("[WARNING]"):
                    if "Could not fetch resource" in line:
                        m = re.search(r"resource (.+?):", line)
                        resource = m.group(1) if m else line
                        warnings.append(f"Imagen no incluida: {resource}")
                    else:
                        warnings.append(line.replace("[WARNING] ", ""))

        if result.returncode != 0:
            raise RuntimeError(
                f"Pandoc terminó con error (código {result.returncode}):\n"
                f"{result.stderr}"
            )

        if not output_path.exists():
            raise RuntimeError("Pandoc no generó el archivo de salida.")

        # --- Post-procesar: aplicar Times New Roman y negrita en títulos ---
        _postprocess_docx(output_path)

    finally:
        # Limpiar archivo temporal
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    return ConversionResult(output_path=str(output_path), warnings=warnings)
