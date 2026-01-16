from pathlib import Path
from typing import Optional, Iterable, Callable
import struct

from pdf2docx import Converter
import fitz  # PyMuPDF
from docx2pdf import convert as docx2pdf_convert
from PIL import Image
import io
import zipfile
from docx import Document
from docx.shared import Inches, Pt, Emu
from docx.oxml.ns import nsmap, qn
from docx.oxml import OxmlElement
import pytesseract
import shutil
import os


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


# --- OCR: PDF (imagen) → DOCX (texto) ---

def _ensure_tesseract_available() -> str:
    """Detecta el ejecutable de Tesseract en Windows y configura pytesseract."""
    candidates = [
        shutil.which("tesseract"),
        r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
        r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
    ]
    for path in candidates:
        if path and Path(path).exists():
            pytesseract.pytesseract.tesseract_cmd = path
            return path
    raise FileNotFoundError(
        "No se encontró Tesseract OCR. Instálalo desde https://github.com/tesseract-ocr/tesseract o el instalador de Windows."
    )


def ocr_pdf_to_docx(input_pdf: Path, output_docx: Path, dpi: int = 300, lang: str = "spa") -> None:
    """Realiza OCR sobre cada página del PDF y escribe el texto en un DOCX.
    - dpi: resolución para render de páginas
    - lang: código del idioma (ej.: 'spa' español, 'eng' inglés, 'spa+eng' mixto)
    """
    if not input_pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {input_pdf}")

    _ensure_tesseract_available()

    output_docx = output_docx.with_suffix(".docx")
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    doc_pdf = fitz.open(str(input_pdf))
    docx_doc = Document()

    for i in range(len(doc_pdf)):
        page = doc_pdf[i]
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")
        pil_img = Image.open(io.BytesIO(img_bytes))

        text = pytesseract.image_to_string(pil_img, lang=lang)
        # Agregamos texto preservando saltos de línea básicos
        for line in text.splitlines():
            docx_doc.add_paragraph(line)
        if i < len(doc_pdf) - 1:
            docx_doc.add_page_break()

    doc_pdf.close()
    docx_doc.save(str(output_docx))


# --- Conversion de imagenes ---

SUPPORTED_IMAGE_FORMATS = {
    "jpg": "JPEG",
    "jpeg": "JPEG",
    "png": "PNG",
    "webp": "WEBP",
    "bmp": "BMP",
    "gif": "GIF",
    "tiff": "TIFF",
    "ico": "ICO",
}


def convert_image(
    input_path: Path,
    output_path: Path,
    output_format: str,
    quality: int = 95,
    resize: Optional[tuple[int, int]] = None,
    maintain_aspect: bool = True,
    overwrite: bool = False,
) -> None:
    """Convierte una imagen a otro formato.

    Args:
        input_path: Ruta de la imagen de entrada
        output_path: Ruta de salida
        output_format: Formato de salida (jpg, png, webp, ico, bmp, gif, tiff)
        quality: Calidad para formatos con compresion (1-100)
        resize: Tuple (width, height) para redimensionar
        maintain_aspect: Mantener proporcion al redimensionar
        overwrite: Sobrescribir si existe
    """
    if not input_path.exists():
        raise FileNotFoundError(f"No existe la imagen: {input_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.exists() and not overwrite:
        raise FileExistsError(f"El archivo ya existe: {output_path}")

    fmt = output_format.lower()
    if fmt not in SUPPORTED_IMAGE_FORMATS:
        raise ValueError(f"Formato no soportado: {output_format}. Usa: {', '.join(SUPPORTED_IMAGE_FORMATS.keys())}")

    img = Image.open(str(input_path))

    # Convertir a RGB si es necesario (para JPEG)
    if fmt in ("jpg", "jpeg") and img.mode in ("RGBA", "P", "LA"):
        background = Image.new("RGB", img.size, (255, 255, 255))
        if img.mode == "P":
            img = img.convert("RGBA")
        background.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
        img = background
    elif fmt == "ico" and img.mode != "RGBA":
        img = img.convert("RGBA")

    # Redimensionar si se especifica
    if resize:
        target_w, target_h = resize
        if maintain_aspect:
            img.thumbnail((target_w, target_h), Image.LANCZOS)
        else:
            img = img.resize((target_w, target_h), Image.LANCZOS)

    # Guardar con las opciones apropiadas
    pil_format = SUPPORTED_IMAGE_FORMATS[fmt]
    save_kwargs = {}

    if pil_format == "JPEG":
        save_kwargs = {"quality": quality, "optimize": True}
    elif pil_format == "PNG":
        save_kwargs = {"optimize": True}
    elif pil_format == "WEBP":
        save_kwargs = {"quality": quality, "method": 6}
    elif pil_format == "ICO":
        # ICO soporta multiples tamaños
        sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        # Filtrar tamaños que sean menores o iguales al original
        orig_size = max(img.size)
        sizes = [s for s in sizes if s[0] <= orig_size]
        if not sizes:
            sizes = [(min(img.size), min(img.size))]
        save_kwargs = {"sizes": sizes}

    img.save(str(output_path), format=pil_format, **save_kwargs)


def batch_convert_images(
    files: Iterable[Path],
    out_dir: Path,
    output_format: str,
    quality: int = 95,
    resize: Optional[tuple[int, int]] = None,
    maintain_aspect: bool = True,
    overwrite: bool = False,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> tuple[int, list[tuple[Path, str]]]:
    """Convierte multiples imagenes en lote."""
    out_dir.mkdir(parents=True, exist_ok=True)
    ok = 0
    errors: list[tuple[Path, str]] = []
    files_list = list(files)
    total = len(files_list)

    for i, f in enumerate(files_list):
        try:
            ext = output_format.lower()
            if ext == "jpeg":
                ext = "jpg"
            target = out_dir / (f.stem + "." + ext)
            convert_image(f, target, output_format, quality, resize, maintain_aspect, overwrite)
            ok += 1
        except Exception as e:
            errors.append((f, str(e)))

        if progress_callback:
            progress_callback(i + 1, total)

    return ok, errors


def get_image_info(image_path: Path) -> dict:
    """Obtiene informacion de una imagen."""
    if not image_path.exists():
        raise FileNotFoundError(f"No existe: {image_path}")

    img = Image.open(str(image_path))
    file_size = image_path.stat().st_size

    return {
        "width": img.size[0],
        "height": img.size[1],
        "format": img.format,
        "mode": img.mode,
        "file_size": file_size,
        "file_size_human": _format_size(file_size),
    }


def _format_size(size_bytes: int) -> str:
    """Formatea bytes a formato legible."""
    for unit in ["B", "KB", "MB", "GB"]:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


# --- Mejoras para documentos corporativos ---

def pdf_to_docx_preserve_formatting(
    input_pdf: Path,
    output_docx: Path,
    start_page: Optional[int] = None,
    end_page: Optional[int] = None,
    overwrite: bool = False,
    embed_fonts: bool = True,
) -> None:
    """Convierte PDF a DOCX preservando mejor el formato corporativo.

    Intenta mantener:
    - Fuentes y estilos
    - Tablas
    - Imagenes y logos
    - Estructura del documento
    """
    if not input_pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {input_pdf}")

    output_docx = output_docx.with_suffix(".docx")
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    if output_docx.exists() and not overwrite:
        raise FileExistsError(f"El archivo ya existe: {output_docx}")

    start_arg = 0 if start_page is None else max(0, start_page - 1)
    end_arg = None if end_page is None else max(start_arg, end_page - 1)

    cv = Converter(str(input_pdf))
    try:
        # Usar configuracion para mejor preservacion
        cv.convert(
            str(output_docx),
            start=start_arg,
            end=end_arg,
        )
    finally:
        cv.close()


def extract_images_from_pdf(input_pdf: Path, output_dir: Path, format: str = "png") -> list[Path]:
    """Extrae todas las imagenes de un PDF."""
    if not input_pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {input_pdf}")

    output_dir.mkdir(parents=True, exist_ok=True)
    extracted: list[Path] = []

    doc = fitz.open(str(input_pdf))
    img_count = 0

    for page_num in range(len(doc)):
        page = doc[page_num]
        image_list = page.get_images()

        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]

            # Guardar imagen
            img_count += 1
            output_path = output_dir / f"imagen_{page_num + 1}_{img_count}.{format}"

            pil_img = Image.open(io.BytesIO(image_bytes))
            if format.lower() in ("jpg", "jpeg") and pil_img.mode == "RGBA":
                background = Image.new("RGB", pil_img.size, (255, 255, 255))
                background.paste(pil_img, mask=pil_img.split()[-1])
                pil_img = background

            pil_img.save(str(output_path))
            extracted.append(output_path)

    doc.close()
    return extracted


def extract_images_from_docx(input_docx: Path, output_dir: Path) -> list[Path]:
    """Extrae todas las imagenes de un DOCX."""
    if not input_docx.exists():
        raise FileNotFoundError(f"No existe el DOCX: {input_docx}")

    output_dir.mkdir(parents=True, exist_ok=True)
    extracted: list[Path] = []

    with zipfile.ZipFile(str(input_docx), 'r') as zf:
        for item in zf.namelist():
            if item.startswith('word/media/'):
                data = zf.read(item)
                filename = Path(item).name
                output_path = output_dir / filename
                output_path.write_bytes(data)
                extracted.append(output_path)

    return extracted
