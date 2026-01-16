from pathlib import Path
from typing import Optional

from pdf2docx import Converter
import fitz  # PyMuPDF
from docx2pdf import convert as docx2pdf_convert
from PIL import Image
import io
import zipfile
from docx import Document
from docx.shared import Inches
from typing import Iterable


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
 

def pdf_to_docx_raster(input_pdf: Path, output_docx: Path, dpi: int = 200, overwrite: bool = False) -> None:
    """Convierte cada página del PDF a imagen y la inserta en un DOCX.
    Máxima fidelidad visual (no editable)."""
    if not input_pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {input_pdf}")

    output_docx = output_docx.with_suffix(".docx")
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    if output_docx.exists():
        if overwrite:
            try:
                output_docx.unlink()
            except Exception:
                # Si no podemos borrar, lanzamos error explícito
                raise PermissionError(f"No se puede sobrescribir: {output_docx}")
        else:
            raise FileExistsError(f"El archivo de salida ya existe: {output_docx}")

    doc_pdf = fitz.open(str(input_pdf))
    docx_doc = Document()

    # Usar ancho de página menos márgenes para ajustar imagen
    section = docx_doc.sections[0]
    page_width = section.page_width
    left_margin = section.left_margin
    right_margin = section.right_margin
    usable_width = page_width - left_margin - right_margin

    for i in range(len(doc_pdf)):
        page = doc_pdf[i]
        # Escala dpi/72
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")
        stream = io.BytesIO(img_bytes)
        pic = docx_doc.add_picture(stream)
        # Ajustar al ancho usable manteniendo proporción
        pic.width = usable_width
        if i < len(doc_pdf) - 1:
            docx_doc.add_page_break()

    doc_pdf.close()
    docx_doc.save(str(output_docx))


# --- Conversión por lotes ---

def batch_pdf_to_docx(
    files: Iterable[Path],
    out_dir: Path,
    mode: str = "editable",  # "editable" | "raster"
    start_page: Optional[int] = None,
    end_page: Optional[int] = None,
    overwrite: bool = False,
    dpi: int = 200,
) -> tuple[int, list[tuple[Path, str]]]:
    out_dir.mkdir(parents=True, exist_ok=True)
    ok = 0
    errors: list[tuple[Path, str]] = []
    for f in files:
        if f.suffix.lower() != ".pdf":
            continue
        target = out_dir / (f.stem + ".docx")
        try:
            if mode == "raster":
                pdf_to_docx_raster(f, target, dpi=dpi, overwrite=overwrite)
            else:
                pdf_to_docx(f, target, start_page, end_page, overwrite)
            ok += 1
        except Exception as e:
            errors.append((f, str(e)))
    return ok, errors


def batch_docx_to_pdf(
    files: Iterable[Path],
    out_dir: Path,
    overwrite: bool = False,
) -> tuple[int, list[tuple[Path, str]]]:
    out_dir.mkdir(parents=True, exist_ok=True)
    ok = 0
    errors: list[tuple[Path, str]] = []
    for f in files:
        if f.suffix.lower() != ".docx":
            continue
        target = out_dir / (f.stem + ".pdf")
        try:
            docx_to_pdf(f, target, overwrite)
            ok += 1
        except Exception as e:
            errors.append((f, str(e)))
    return ok, errors


def scan_files(directory: Path) -> tuple[list[Path], list[Path]]:
    """Devuelve listas de PDFs y DOCXs en la carpeta (no recursivo)."""
    if not directory.exists() or not directory.is_dir():
        raise FileNotFoundError(f"No existe la carpeta: {directory}")
    pdfs = sorted(directory.glob("*.pdf"))
    docxs = sorted(directory.glob("*.docx"))
    return pdfs, docxs
 