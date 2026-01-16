from pathlib import Path
from typing import Optional

from pdf2docx import Converter
import fitz  # PyMuPDF
from docx2pdf import convert as docx2pdf_convert
from PIL import Image
import io
import zipfile


def pdf_to_docx(input_pdf: Path, output_docx: Path, start_page: Optional[int], end_page: Optional[int], overwrite: bool) -> None:
    if not input_pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {input_pdf}")

    output_docx = output_docx.with_suffix(".docx")
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    if output_docx.exists() and not overwrite:
        raise FileExistsError(f"El archivo de salida ya existe: {output_docx}. Usa --overwrite para reemplazarlo.")

    start_arg = 0 if start_page is None else max(0, start_page - 1)
    end_arg = None if end_page is None else max(start_arg, end_page - 1)

    cv = Converter(str(input_pdf))
    try:
        cv.convert(str(output_docx), start=start_arg, end=end_arg)
    finally:
        cv.close()


def docx_to_pdf(input_docx: Path, output_pdf: Path, overwrite: bool) -> None:
    if not input_docx.exists():
        raise FileNotFoundError(f"No existe el DOCX: {input_docx}")

    output_pdf = output_pdf.with_suffix(".pdf")
    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    if output_pdf.exists() and not overwrite:
        raise FileExistsError(f"El archivo de salida ya existe: {output_pdf}. Usa --overwrite para reemplazarlo.")

    # Usa Microsoft Word (COM) en Windows si está disponible; en macOS usa automatización.
    # Si Word no está instalado, lanzará una excepción.
    docx2pdf_convert(str(input_docx), str(output_pdf))


def compress_pdf(input_pdf: Path, output_pdf: Path) -> None:
    if not input_pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {input_pdf}")

    doc = fitz.open(str(input_pdf))
    # Garbage=4 limpia objetos no usados; deflate=True comprime streams; linear=False para tamaño
    doc.save(str(output_pdf), garbage=4, deflate=True)
    doc.close()


def _resize_image(img: Image.Image, max_width: Optional[int], max_height: Optional[int]) -> Image.Image:
    if max_width is None and max_height is None:
        return img
    w, h = img.size
    ratio_w = max_width / w if max_width else 1.0
    ratio_h = max_height / h if max_height else 1.0
    ratio = min(ratio_w, ratio_h)
    if ratio < 1.0:
        new_size = (max(1, int(w * ratio)), max(1, int(h * ratio)))
        return img.resize(new_size, Image.LANCZOS)
    return img


def compress_docx_images(input_docx: Path, output_docx: Path, quality: int = 75, max_width: Optional[int] = None, max_height: Optional[int] = None) -> None:
    if not input_docx.exists():
        raise FileNotFoundError(f"No existe el DOCX: {input_docx}")

    output_docx = output_docx.with_suffix(".docx")
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(str(input_docx), 'r') as zin, zipfile.ZipFile(str(output_docx), 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.lower().startswith('word/media/'):
                # Intentar recomprimir imagen
                try:
                    img = Image.open(io.BytesIO(data))
                    img = _resize_image(img, max_width, max_height)
                    buf = io.BytesIO()
                    fmt = (img.format or '').upper()
                    if fmt in ('JPEG', 'JPG'):
                        img.save(buf, format='JPEG', quality=quality, optimize=True)
                        data = buf.getvalue()
                    elif fmt in ('PNG',):
                        # PNG: optimizar sin perder; para compresión mayor, podría convertirse a JPEG
                        img.save(buf, format='PNG', optimize=True)
                        data = buf.getvalue()
                    else:
                        # Otros formatos: intentar guardar en mismo formato si posible
                        save_fmt = fmt if fmt else 'PNG'
                        img.save(buf, format=save_fmt, optimize=True)
                        data = buf.getvalue()
                except Exception:
                    # Si no se pudo, dejar el original
                    pass
            zout.writestr(item, data)
*** End Patch