"""
Wrapper sobre la librería markitdown de Microsoft.
Convierte cualquier archivo soportado a formato Markdown.

No tiene código de UI — lógica pura para usar desde ConversionWorker.
"""
from pathlib import Path


def convert_to_markdown(input_path: str, output_path: str) -> str:
    """
    Convierte el archivo en input_path a Markdown y lo guarda en output_path.

    Args:
        input_path:  Ruta absoluta del archivo de entrada.
        output_path: Ruta absoluta donde se guardará el .md generado.

    Returns:
        El contenido Markdown como string.

    Raises:
        Exception: Cualquier error de markitdown se deja propagar
                   para que el caller (ConversionWorker) lo capture.
    """
    from markitdown import MarkItDown

    md = MarkItDown()
    result = md.convert(input_path)
    content = result.text_content

    # Asegurar que el directorio de salida existe
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    Path(output_path).write_text(content, encoding="utf-8")

    return content
