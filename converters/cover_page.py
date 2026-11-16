r"""
Generación de carátulas (portadas) para documentos Word.

Reproduce la plantilla institucional entregada:
  - Universidad / Facultad / Escuela centradas y en negrita.
  - Logo institucional centrado.
  - Título del trabajo.
  - Tabla de Integrantes, Docente, Asignatura y Semestre Académico.
  - Ciudad y año.

Los datos se persisten en un archivo JSON dentro de la carpeta de
configuración del usuario, de modo que no haya que escribirlos de nuevo.
"""
from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path


_FONT_NAME = "Times New Roman"
_APP_DIR_NAME = "MarkdownConverter"
_CONFIG_NAME = "caratula_config.json"


@dataclass
class CoverData:
    """Datos editables de la carátula."""

    institution: str = ""
    faculty: str = ""
    school: str = ""
    title: str = ""
    students: list[str] = field(default_factory=list)
    teacher: str = ""
    subject: str = ""
    semester: str = ""
    city: str = ""
    year: str = ""
    logo_path: str = ""

    def is_empty(self) -> bool:
        """True si no hay ningún dato que justifique generar una carátula."""
        return not any(
            [
                self.institution,
                self.faculty,
                self.school,
                self.title,
                self.teacher,
                self.subject,
                self.semester,
                self.city,
                self.year,
                self.logo_path,
                self.students,
            ]
        )


# ---------------------------------------------------------------------------
# Persistencia de la configuración
# ---------------------------------------------------------------------------

def config_path() -> Path:
    """Ruta del archivo JSON de configuración de la carátula."""
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home())
        return base / _APP_DIR_NAME / _CONFIG_NAME
    base = Path(os.environ.get("XDG_CONFIG_HOME") or (Path.home() / ".config"))
    return base / "markdown-converter" / _CONFIG_NAME


def load_cover_config() -> tuple[bool, CoverData]:
    """
    Lee la configuración guardada.

    Returns:
        (habilitada, datos). Si no existe o está corrupta, devuelve
        (False, CoverData()) con el año actual como valor por defecto.
    """
    default_year = str(datetime.now().year)
    path = config_path()
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False, CoverData(year=default_year)

    enabled = bool(raw.get("enabled", False))
    cover_raw = raw.get("cover", {})
    if not isinstance(cover_raw, dict):
        cover_raw = {}

    students = cover_raw.get("students", [])
    if isinstance(students, str):
        students = [line.strip() for line in students.splitlines() if line.strip()]
    elif isinstance(students, list):
        students = [str(s).strip() for s in students if str(s).strip()]
    else:
        students = []

    data = CoverData(
        institution=str(cover_raw.get("institution", "") or ""),
        faculty=str(cover_raw.get("faculty", "") or ""),
        school=str(cover_raw.get("school", "") or ""),
        title=str(cover_raw.get("title", "") or ""),
        students=students,
        teacher=str(cover_raw.get("teacher", "") or ""),
        subject=str(cover_raw.get("subject", "") or ""),
        semester=str(cover_raw.get("semester", "") or ""),
        city=str(cover_raw.get("city", "") or ""),
        year=str(cover_raw.get("year", "") or default_year),
        logo_path=str(cover_raw.get("logo_path", "") or ""),
    )
    return enabled, data


def save_cover_config(enabled: bool, data: CoverData) -> None:
    """Guarda la configuración de forma atómica."""
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"version": 1, "enabled": bool(enabled), "cover": asdict(data)}
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


# ---------------------------------------------------------------------------
# Helpers OpenXML
# ---------------------------------------------------------------------------

def _style_run(run, size_pt: float, bold: bool) -> None:
    """Fuerza Times New Roman, tamaño y color negro en un run."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    run.bold = bold
    run.font.size = Pt(size_pt)
    run.font.name = _FONT_NAME
    run.font.color.rgb = RGBColor(0, 0, 0)

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), _FONT_NAME)


def _format_paragraph(para, before: float | None, after: float | None, line: float | None) -> None:
    """Aplica espaciado a un párrafo."""
    from docx.shared import Pt

    pf = para.paragraph_format
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line


def _set_cell_borders(cell) -> None:
    """Borde sutil gris claro, como en la plantilla."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:tcBorders"))
    if old is not None:
        tcPr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "E2E8F0")
        borders.append(el)
    tcPr.insert_element_before(
        borders,
        "w:shd",
        "w:noWrap",
        "w:tcMar",
        "w:textDirection",
        "w:tcFitText",
        "w:vAlign",
        "w:hideMark",
    )


def _set_cell_shading(cell, fill: str) -> None:
    """Fondo de celda (la primera columna de la plantilla usa gris muy claro)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:shd"))
    if old is not None:
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), fill)
    tcPr.insert_element_before(
        shd,
        "w:noWrap",
        "w:tcMar",
        "w:textDirection",
        "w:tcFitText",
        "w:vAlign",
        "w:hideMark",
    )


def _set_cell_margins(cell, value_twips: int = 90) -> None:
    """Márgenes internos de la celda."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:tcMar"))
    if old is not None:
        tcPr.remove(old)
    mar = OxmlElement("w:tcMar")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(value_twips))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.insert_element_before(
        mar,
        "w:textDirection",
        "w:tcFitText",
        "w:vAlign",
        "w:hideMark",
    )


def _set_table_borders(table) -> None:
    """Bordes de tabla gris claro e interior, como en la plantilla."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "E2E8F0")
        borders.append(el)
    tblPr.insert_element_before(
        borders,
        "w:shd",
        "w:tblLayout",
        "w:tblCellMar",
        "w:tblLook",
        "w:tblCaption",
        "w:tblDescription",
    )


def _set_table_width(table, col_twips: list[int]) -> None:
    """Fija el ancho de columnas con layout fijo (dxa)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    total = int(sum(col_twips))

    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.insert_element_before(
            tblW,
            "w:jc",
            "w:tblCellSpacing",
            "w:tblInd",
            "w:tblBorders",
            "w:shd",
            "w:tblLayout",
            "w:tblCellMar",
            "w:tblLook",
        )
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")

    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.insert_element_before(
            layout,
            "w:tblCellMar",
            "w:tblLook",
            "w:tblCaption",
            "w:tblDescription",
        )
    layout.set(qn("w:type"), "fixed")

    for grid_col, width in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), col_twips):
        grid_col.set(qn("w:w"), str(width))

    for row in table.rows:
        for cell, width in zip(row.cells, col_twips):
            cell.width = width


def _fill_cell(cell, lines: list[str], bold: bool, size_pt: float = 12) -> None:
    """Escribe una o varias líneas dentro de una celda."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    for index, line in enumerate(lines):
        para = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if bold:
            _format_paragraph(para, before=4.5, after=4.5, line=1.25)
        else:
            _format_paragraph(para, before=0, after=0, line=1.25)
        run = para.add_run(line)
        _style_run(run, size_pt, bold)


def _add_cover_table(doc, rows: list[tuple[str, list[str]]]) -> object:
    """Crea la tabla de datos de la carátula."""
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT

    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    col_twips = [2621, 4867]  # 1.82" + 3.38" = 5.2" aprox.
    _set_table_width(table, col_twips)
    _set_table_borders(table)

    for row_index, (label, values) in enumerate(rows):
        label_cell, value_cell = table.rows[row_index].cells
        for cell in (label_cell, value_cell):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            _set_cell_borders(cell)
            _set_cell_margins(cell)

        _set_cell_shading(label_cell, "F7FAFC")
        _fill_cell(label_cell, [label], bold=True)
        _fill_cell(value_cell, values or [""], bold=False)

    return table


# ---------------------------------------------------------------------------
# Inserción de la carátula en el .docx
# ---------------------------------------------------------------------------

def insert_cover_page(docx_path: str | Path, cover: CoverData) -> list[str]:
    """
    Inserta la carátula al inicio del .docx y añade un salto de página,
    de modo que el contenido original comience en la página 2.

    Returns:
        Lista de avisos (por ejemplo, si el logo no se pudo incrustar).
    """
    warnings: list[str] = []
    if cover is None or cover.is_empty():
        return warnings

    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.shared import Inches, Mm

    doc = Document(str(docx_path))
    body = doc.element.body
    if len(body) == 0:
        return warnings

    # La plantilla institucional está en A4 con márgenes de 1 pulgada.
    # Se aplica al documento completo para que portada y contenido coincidan.
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.header_distance = Inches(0.5)
        section.footer_distance = Inches(0.5)

    first_element = body[0]
    created: list = []

    def add_paragraph(
        text: str = "",
        *,
        bold: bool = False,
        size_pt: float = 12,
        align=None,
        before: float | None = None,
        after: float | None = None,
        line: float | None = 1.25,
    ):
        para = doc.add_paragraph()
        if align is not None:
            para.alignment = align
        _format_paragraph(para, before, after, line)
        if text:
            run = para.add_run(text)
            _style_run(run, size_pt, bold)
        created.append(para._p)
        return para

    center = WD_ALIGN_PARAGRAPH.CENTER

    if cover.institution:
        add_paragraph(
            cover.institution,
            bold=True,
            size_pt=22,
            align=center,
            before=12,
            after=3.75,
        )

    if cover.faculty:
        add_paragraph(
            cover.faculty,
            bold=True,
            size_pt=15,
            align=center,
            before=0,
            after=3.75,
        )

    if cover.school:
        add_paragraph(
            cover.school,
            bold=True,
            size_pt=13,
            align=center,
            before=0,
            after=6,
        )

    if cover.logo_path:
        logo = Path(cover.logo_path)
        if not logo.exists():
            warnings.append(f"Logo de carátula no encontrado: {cover.logo_path}")
        else:
            logo_para = doc.add_paragraph()
            logo_para.alignment = center
            _format_paragraph(logo_para, before=0, after=6, line=1.0)
            run = logo_para.add_run()
            section = doc.sections[0]
            page_width = section.page_width or Inches(8.5)
            left_margin = section.left_margin or Inches(1)
            right_margin = section.right_margin or Inches(1)
            max_width = page_width - left_margin - right_margin
            width = min(Inches(4.5), max_width)

            # Evita que un logo muy alto desborde la primera página.
            try:
                from docx.image.image import Image as DocxImage

                image = DocxImage.from_file(str(logo))
                if image.width and image.height:
                    max_height = Inches(2.7)
                    height = int(width * image.height / image.width)
                    if height > max_height:
                        width = int(max_height * image.width / image.height)
            except Exception:
                pass

            try:
                run.add_picture(str(logo), width=width)
                created.append(logo_para._p)
            except Exception:
                logo_para._p.getparent().remove(logo_para._p)
                warnings.append(
                    f"No se pudo incrustar el logo de la carátula: {cover.logo_path}"
                )

    add_paragraph("", before=0, after=6, line=1.0)

    if cover.title:
        add_paragraph(
            cover.title,
            bold=True,
            size_pt=18,
            align=center,
            before=24,
            after=0,
        )

    add_paragraph("", before=0, after=18.75)

    rows: list[tuple[str, list[str]]] = []
    if cover.students:
        rows.append(("Integrantes:", cover.students))
    if cover.teacher:
        rows.append(("Docente:", [cover.teacher]))
    if cover.subject:
        rows.append(("Asignatura:", [cover.subject]))
    if cover.semester:
        rows.append(("Semestre Académico:", [cover.semester]))

    if rows:
        table = _add_cover_table(doc, rows)
        created.append(table._tbl)
        add_paragraph("", before=0, after=0)

    if cover.city:
        add_paragraph(
            cover.city,
            bold=True,
            size_pt=12,
            align=center,
            before=18.75,
            after=0,
        )

    if cover.year:
        add_paragraph(
            cover.year,
            bold=True,
            size_pt=12,
            align=center,
            before=18.75,
            after=0,
        )

    page_break = doc.add_paragraph()
    _format_paragraph(page_break, before=0, after=0, line=1.0)
    page_break.add_run().add_break(WD_BREAK.PAGE)
    created.append(page_break._p)

    for element in created:
        first_element.addprevious(element)

    doc.save(str(docx_path))
    return warnings
