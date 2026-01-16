"""
Funciones de conversión y procesamiento para PDF Converter Pro.
"""
from pathlib import Path
from typing import Optional, Callable, Any
import os
import io
import shutil
import tempfile
import zipfile

# Type alias para callbacks de progreso
ProgressCallback = Callable[[int, int, str], None]
CancelCheck = Callable[[], bool]

# Formatos de imagen soportados
SUPPORTED_IMAGE_FORMATS = ["png", "jpg", "jpeg", "webp", "bmp", "gif", "tiff", "ico"]


# ===========================================================================
# PDF -> DOCX (Editable)
# ===========================================================================

def pdf_to_docx(
    input_pdf: Path,
    output_docx: Path,
    start: Optional[int] = None,
    end: Optional[int] = None,
    overwrite: bool = False
) -> None:
    """Convierte PDF a DOCX usando pdf2docx (texto editable)."""
    if output_docx.exists() and not overwrite:
        raise FileExistsError(f"El archivo ya existe: {output_docx}")

    from pdf2docx import Converter
    import fitz

    # Normalizar rango: evitar pasar None a la librería
    # Consideramos que 'start' y 'end' vienen 1-basados desde la UI.
    doc = fitz.open(str(input_pdf))
    try:
        total_pages = doc.page_count
    finally:
        doc.close()

    actual_start = (start - 1) if (start and start > 0) else 0
    actual_end = end if (end and end > 0) else total_pages

    cv = Converter(str(input_pdf))
    try:
        cv.convert(str(output_docx), start=actual_start, end=actual_end)
    finally:
        cv.close()


def pdf_to_docx_with_progress(
    input_pdf: Path,
    output_docx: Path,
    start: Optional[int] = None,
    end: Optional[int] = None,
    overwrite: bool = False,
    progress_callback: Optional[ProgressCallback] = None,
    cancel_check: Optional[CancelCheck] = None
) -> None:
    """Convierte PDF a DOCX con reporte de progreso."""
    if output_docx.exists() and not overwrite:
        raise FileExistsError(f"El archivo ya existe: {output_docx}")

    from pdf2docx import Converter
    import fitz

    cv = Converter(str(input_pdf))
    try:
        # Obtener número de páginas
        doc = fitz.open(str(input_pdf))
        total_pages = doc.page_count
        doc.close()

        # Normalizar rango (UI 1-basada → índice 0-basado)
        actual_start = (start - 1) if (start and start > 0) else 0
        actual_end = end if (end and end > 0) else total_pages

        if progress_callback:
            progress_callback(0, actual_end - actual_start, f"Iniciando conversión de {total_pages} páginas...")

        # Convertir página por página para reportar progreso
        for i in range(actual_start, actual_end):
            if cancel_check and cancel_check():
                raise InterruptedError("Operación cancelada por el usuario")

            page_num = i + 1
            if progress_callback:
                progress_callback(i - actual_start + 1, actual_end - actual_start, f"Página {page_num}/{actual_end}")

        # Hacer la conversión real con rango normalizado
        cv.convert(str(output_docx), start=actual_start, end=actual_end)

        if progress_callback:
            progress_callback(actual_end - actual_start, actual_end - actual_start, "Conversión completada")

    finally:
        cv.close()


# ===========================================================================
# DOCX -> PDF
# ===========================================================================

def docx_to_pdf(
    input_docx: Path,
    output_pdf: Path,
    overwrite: bool = False
) -> None:
    """Convierte DOCX a PDF con múltiples métodos de fallback."""
    if output_pdf.exists() and not overwrite:
        raise FileExistsError(f"El archivo ya existe: {output_pdf}")

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    errors = []

    # Método 1: Usar docx2pdf con COM de Word (más preciso)
    try:
        from docx2pdf import convert
        convert(str(input_docx), str(output_pdf))
        if output_pdf.exists():
            return
    except Exception as e:
        errors.append(f"docx2pdf: {e}")

    # Método 2: Usar COM de Word directamente con mejor manejo
    try:
        _docx_to_pdf_word_com(input_docx, output_pdf)
        if output_pdf.exists():
            return
    except Exception as e:
        errors.append(f"Word COM: {e}")

    # Método 3: Renderizar DOCX a imágenes y crear PDF
    try:
        _docx_to_pdf_via_images(input_docx, output_pdf)
        if output_pdf.exists():
            return
    except Exception as e:
        errors.append(f"Via imágenes: {e}")

    raise RuntimeError(f"No se pudo convertir el archivo. Errores: {'; '.join(errors)}")


def _docx_to_pdf_word_com(input_docx: Path, output_pdf: Path) -> None:
    """Convierte DOCX a PDF usando COM de Word directamente."""
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        # Abrir documento
        doc = word.Documents.Open(
            str(input_docx.resolve()),
            ReadOnly=True,
            AddToRecentFiles=False
        )

        # Exportar como PDF (17 = wdFormatPDF)
        doc.SaveAs2(
            str(output_pdf.resolve()),
            FileFormat=17
        )

    finally:
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _docx_to_pdf_via_images(input_docx: Path, output_pdf: Path) -> None:
    """Convierte DOCX a PDF extrayendo todo el contenido (texto, tablas, imágenes)."""
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from PIL import Image
    import fitz
    from io import BytesIO

    # Leer el DOCX
    doc = Document(str(input_docx))

    # Crear PDF
    pdf_doc = fitz.open()

    # Configurar página tamaño carta
    page_width = 612  # 8.5 pulgadas
    page_height = 792  # 11 pulgadas
    margin = 72  # 1 pulgada

    current_page = pdf_doc.new_page(width=page_width, height=page_height)
    y_position = margin
    line_height = 14
    max_y = page_height - margin
    content_width = page_width - 2 * margin

    def new_page_if_needed(space_needed: float) -> None:
        nonlocal current_page, y_position
        if y_position + space_needed > max_y:
            current_page = pdf_doc.new_page(width=page_width, height=page_height)
            y_position = margin

    def insert_text_wrapped(text: str, fontsize: int = 11, bold: bool = False) -> None:
        nonlocal y_position
        if not text.strip():
            y_position += line_height / 2
            return

        fontname = "helv" if not bold else "hebo"
        char_width = fontsize * 0.5  # Aproximación

        words = text.split()
        line = ""
        for word in words:
            test_line = f"{line} {word}".strip()
            if len(test_line) * char_width < content_width:
                line = test_line
            else:
                if line:
                    new_page_if_needed(line_height)
                    current_page.insert_text(
                        (margin, y_position + line_height),
                        line,
                        fontsize=fontsize,
                        fontname=fontname
                    )
                    y_position += line_height
                line = word

        if line:
            new_page_if_needed(line_height)
            current_page.insert_text(
                (margin, y_position + line_height),
                line,
                fontsize=fontsize,
                fontname=fontname
            )
            y_position += line_height * 1.3

    def insert_table(table: Table) -> None:
        nonlocal y_position
        # Calcular dimensiones de la tabla
        num_cols = len(table.columns)
        col_width = content_width / num_cols if num_cols > 0 else content_width
        row_height = line_height * 1.5

        for row in table.rows:
            new_page_if_needed(row_height + 10)

            x_pos = margin
            for cell in row.cells:
                cell_text = cell.text.strip()[:50]  # Limitar texto
                if cell_text:
                    # Dibujar borde de celda
                    rect = fitz.Rect(x_pos, y_position, x_pos + col_width, y_position + row_height)
                    current_page.draw_rect(rect, color=(0.7, 0.7, 0.7), width=0.5)
                    # Insertar texto
                    current_page.insert_text(
                        (x_pos + 3, y_position + line_height),
                        cell_text,
                        fontsize=9,
                        fontname="helv"
                    )
                x_pos += col_width

            y_position += row_height

        y_position += line_height  # Espacio después de tabla

    # Procesar todo el contenido del documento en orden
    for element in doc.element.body:
        tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag

        if tag == 'p':  # Párrafo
            # Buscar el párrafo correspondiente
            for para in doc.paragraphs:
                if para._element == element:
                    text = para.text
                    # Detectar si es título/encabezado
                    is_heading = para.style and para.style.name and 'Heading' in para.style.name
                    insert_text_wrapped(text, fontsize=14 if is_heading else 11, bold=is_heading)
                    break

        elif tag == 'tbl':  # Tabla
            for table in doc.tables:
                if table._element == element:
                    insert_table(table)
                    break

    # Extraer e insertar imágenes del documento
    try:
        with zipfile.ZipFile(str(input_docx), 'r') as zf:
            for name in zf.namelist():
                if name.startswith('word/media/'):
                    try:
                        img_data = zf.read(name)
                        img = Image.open(BytesIO(img_data))

                        # Escalar imagen si es muy grande
                        max_img_width = content_width
                        max_img_height = 300
                        img_w, img_h = img.size

                        # Convertir pixeles a puntos (72 dpi)
                        scale = min(max_img_width / img_w, max_img_height / img_h, 1.0)
                        display_w = img_w * scale
                        display_h = img_h * scale

                        new_page_if_needed(display_h + 20)

                        # Insertar imagen
                        img_rect = fitz.Rect(
                            margin, y_position,
                            margin + display_w, y_position + display_h
                        )
                        img_bytes = BytesIO()
                        if img.mode in ('RGBA', 'P'):
                            img = img.convert('RGB')
                        img.save(img_bytes, format='PNG')
                        current_page.insert_image(img_rect, stream=img_bytes.getvalue())
                        y_position += display_h + 10

                    except Exception:
                        pass  # Ignorar imágenes que no se puedan procesar
    except Exception:
        pass  # Ignorar errores al extraer imágenes

    pdf_doc.save(str(output_pdf))
    pdf_doc.close()


def docx_to_pdf_with_progress(
    input_docx: Path,
    output_pdf: Path,
    overwrite: bool = False,
    progress_callback: Optional[ProgressCallback] = None,
    cancel_check: Optional[CancelCheck] = None
) -> None:
    """Convierte DOCX a PDF con reporte de progreso."""
    if output_pdf.exists() and not overwrite:
        raise FileExistsError(f"El archivo ya existe: {output_pdf}")

    if progress_callback:
        progress_callback(0, 3, "Iniciando conversión...")

    if cancel_check and cancel_check():
        raise InterruptedError("Operación cancelada")

    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    if progress_callback:
        progress_callback(1, 3, "Convirtiendo documento...")

    # Intentar métodos en orden
    errors = []

    # Método 1: docx2pdf
    try:
        if progress_callback:
            progress_callback(1, 3, "Intentando con Microsoft Word...")
        from docx2pdf import convert
        convert(str(input_docx), str(output_pdf))
        if output_pdf.exists():
            if progress_callback:
                progress_callback(3, 3, "Conversión completada")
            return
    except Exception as e:
        errors.append(f"docx2pdf: {e}")

    if cancel_check and cancel_check():
        raise InterruptedError("Operación cancelada")

    # Método 2: COM directo
    try:
        if progress_callback:
            progress_callback(2, 3, "Intentando método alternativo...")
        _docx_to_pdf_word_com(input_docx, output_pdf)
        if output_pdf.exists():
            if progress_callback:
                progress_callback(3, 3, "Conversión completada")
            return
    except Exception as e:
        errors.append(f"Word COM: {e}")

    if cancel_check and cancel_check():
        raise InterruptedError("Operación cancelada")

    # Método 3: Via imágenes
    try:
        if progress_callback:
            progress_callback(2, 3, "Usando conversión de emergencia...")
        _docx_to_pdf_via_images(input_docx, output_pdf)
        if output_pdf.exists():
            if progress_callback:
                progress_callback(3, 3, "Conversión completada (modo básico)")
            return
    except Exception as e:
        errors.append(f"Via imágenes: {e}")

    raise RuntimeError(f"No se pudo convertir. Errores: {'; '.join(errors)}")


# ===========================================================================
# PDF -> DOCX (Raster/Imagen)
# ===========================================================================

def pdf_to_docx_raster(
    input_pdf: Path,
    output_docx: Path,
    dpi: int = 200,
    overwrite: bool = False
) -> None:
    """Convierte PDF a DOCX renderizando como imágenes (fidelidad exacta)."""
    if output_docx.exists() and not overwrite:
        raise FileExistsError(f"El archivo ya existe: {output_docx}")

    import fitz
    from docx import Document
    from docx.shared import Inches

    doc = fitz.open(str(input_pdf))
    word_doc = Document()

    try:
        for page_num in range(doc.page_count):
            page = doc[page_num]
            # Renderizar página como imagen
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)

            # Guardar temporalmente
            img_data = pix.tobytes("png")

            # Insertar en Word
            from io import BytesIO
            img_stream = BytesIO(img_data)

            # Calcular tamaño en pulgadas (basado en tamaño de página)
            width_inches = page.rect.width / 72
            word_doc.add_picture(img_stream, width=Inches(min(width_inches, 7.5)))

            if page_num < doc.page_count - 1:
                word_doc.add_page_break()

        word_doc.save(str(output_docx))
    finally:
        doc.close()


def pdf_to_docx_raster_with_progress(
    input_pdf: Path,
    output_docx: Path,
    dpi: int = 200,
    overwrite: bool = False,
    progress_callback: Optional[ProgressCallback] = None,
    cancel_check: Optional[CancelCheck] = None
) -> None:
    """Convierte PDF a DOCX como imágenes con reporte de progreso."""
    if output_docx.exists() and not overwrite:
        raise FileExistsError(f"El archivo ya existe: {output_docx}")

    import fitz
    from docx import Document
    from docx.shared import Inches
    from io import BytesIO

    doc = fitz.open(str(input_pdf))
    word_doc = Document()
    total_pages = doc.page_count

    if progress_callback:
        progress_callback(0, total_pages, f"Procesando {total_pages} páginas a {dpi} DPI...")

    try:
        for page_num in range(total_pages):
            if cancel_check and cancel_check():
                raise InterruptedError("Operación cancelada por el usuario")

            if progress_callback:
                progress_callback(page_num + 1, total_pages, f"Renderizando página {page_num + 1}/{total_pages}")

            page = doc[page_num]
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img_stream = BytesIO(img_data)

            width_inches = page.rect.width / 72
            word_doc.add_picture(img_stream, width=Inches(min(width_inches, 7.5)))

            if page_num < total_pages - 1:
                word_doc.add_page_break()

        if progress_callback:
            progress_callback(total_pages, total_pages, "Guardando documento...")

        word_doc.save(str(output_docx))

    finally:
        doc.close()


# ===========================================================================
# OCR PDF -> DOCX
# ===========================================================================

def check_tesseract_installed() -> tuple[bool, str]:
    """Verifica si Tesseract OCR está instalado."""
    import shutil

    # Verificar si está en PATH
    tesseract_path = shutil.which("tesseract")
    if tesseract_path:
        return True, tesseract_path

    # Rutas comunes en Windows
    common_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
    ]

    for path in common_paths:
        if os.path.exists(path):
            return True, path

    return False, ""


def ocr_pdf_to_docx_with_progress(
    input_pdf: Path,
    output_docx: Path,
    dpi: int = 300,
    lang: str = "spa",
    progress_callback: Optional[ProgressCallback] = None,
    cancel_check: Optional[CancelCheck] = None
) -> None:
    """Convierte PDF a DOCX usando OCR (pytesseract)."""
    import fitz
    from docx import Document
    from PIL import Image
    import pytesseract
    from io import BytesIO

    # Verificar Tesseract antes de empezar
    tesseract_ok, tesseract_path = check_tesseract_installed()
    if not tesseract_ok:
        raise RuntimeError(
            "Tesseract OCR no está instalado.\n\n"
            "Para usar OCR, instala Tesseract:\n"
            "1. Descarga desde: https://github.com/UB-Mannheim/tesseract/wiki\n"
            "2. Instala y agrega al PATH del sistema\n"
            "3. Reinicia la aplicación"
        )

    # Configurar ruta de Tesseract si se encontró
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    doc = fitz.open(str(input_pdf))
    word_doc = Document()
    total_pages = doc.page_count

    if progress_callback:
        progress_callback(0, total_pages, f"Iniciando OCR ({lang})...")

    try:
        for page_num in range(total_pages):
            if cancel_check and cancel_check():
                raise InterruptedError("Operación cancelada por el usuario")

            if progress_callback:
                progress_callback(page_num + 1, total_pages, f"OCR página {page_num + 1}/{total_pages}")

            page = doc[page_num]
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)

            # Convertir a PIL Image para OCR
            img_data = pix.tobytes("png")
            img = Image.open(BytesIO(img_data))

            # Ejecutar OCR
            text = pytesseract.image_to_string(img, lang=lang)

            # Agregar texto al documento
            if text.strip():
                word_doc.add_paragraph(text)

            if page_num < total_pages - 1:
                word_doc.add_page_break()

        if progress_callback:
            progress_callback(total_pages, total_pages, "Guardando documento...")

        word_doc.save(str(output_docx))

    finally:
        doc.close()


# ===========================================================================
# Compresión PDF
# ===========================================================================

def compress_pdf_with_progress(
    input_pdf: Path,
    output_pdf: Path,
    progress_callback: Optional[ProgressCallback] = None,
    cancel_check: Optional[CancelCheck] = None
) -> dict:
    """Optimiza/comprime un PDF."""
    import pikepdf

    original_size = input_pdf.stat().st_size

    if progress_callback:
        progress_callback(0, 3, "Abriendo PDF...")

    if cancel_check and cancel_check():
        raise InterruptedError("Operación cancelada")

    with pikepdf.open(str(input_pdf)) as pdf:
        if progress_callback:
            progress_callback(1, 3, "Optimizando contenido...")

        if cancel_check and cancel_check():
            raise InterruptedError("Operación cancelada")

        if progress_callback:
            progress_callback(2, 3, "Guardando PDF optimizado...")

        pdf.save(
            str(output_pdf),
            compress_streams=True,
            object_stream_mode=pikepdf.ObjectStreamMode.generate
        )

    new_size = output_pdf.stat().st_size
    reduction = ((original_size - new_size) / original_size) * 100 if original_size > 0 else 0

    if progress_callback:
        progress_callback(3, 3, "Compresión completada")

    return {
        "original_size": original_size,
        "new_size": new_size,
        "reduction_percent": max(0, reduction)
    }


# ===========================================================================
# Compresión imágenes DOCX
# ===========================================================================

def compress_docx_images_with_progress(
    input_docx: Path,
    output_docx: Path,
    quality: int = 75,
    max_width: Optional[int] = None,
    max_height: Optional[int] = None,
    progress_callback: Optional[ProgressCallback] = None,
    cancel_check: Optional[CancelCheck] = None
) -> dict:
    """Comprime las imágenes dentro de un archivo DOCX."""
    from PIL import Image

    original_size = input_docx.stat().st_size

    # Crear directorio temporal
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        extract_dir = tmpdir_path / "extracted"

        if progress_callback:
            progress_callback(0, 4, "Extrayendo DOCX...")

        # Extraer DOCX (es un ZIP)
        with zipfile.ZipFile(str(input_docx), 'r') as zip_ref:
            zip_ref.extractall(str(extract_dir))

        if cancel_check and cancel_check():
            raise InterruptedError("Operación cancelada")

        # Buscar imágenes
        media_dir = extract_dir / "word" / "media"
        images_processed = 0

        if media_dir.exists():
            image_files = list(media_dir.glob("*"))
            total_images = len(image_files)

            if progress_callback:
                progress_callback(1, 4, f"Comprimiendo {total_images} imágenes...")

            for i, img_path in enumerate(image_files):
                if cancel_check and cancel_check():
                    raise InterruptedError("Operación cancelada")

                try:
                    # Abrir imagen
                    with Image.open(img_path) as img:
                        # Convertir a RGB si es necesario
                        if img.mode in ('RGBA', 'P'):
                            img = img.convert('RGB')

                        # Redimensionar si se especificó
                        if max_width or max_height:
                            w, h = img.size
                            new_w, new_h = w, h

                            if max_width and w > max_width:
                                ratio = max_width / w
                                new_w = max_width
                                new_h = int(h * ratio)

                            if max_height and new_h > max_height:
                                ratio = max_height / new_h
                                new_h = max_height
                                new_w = int(new_w * ratio)

                            if new_w != w or new_h != h:
                                img = img.resize((new_w, new_h), Image.LANCZOS)

                        # Guardar como JPEG comprimido
                        new_path = img_path.with_suffix('.jpeg')
                        img.save(str(new_path), 'JPEG', quality=quality, optimize=True)

                        # Eliminar original si es diferente
                        if new_path != img_path:
                            img_path.unlink()

                        images_processed += 1

                except Exception:
                    # Si falla, dejar la imagen original
                    pass

                if progress_callback and total_images > 0:
                    progress_callback(1, 4, f"Imagen {i + 1}/{total_images}")

        if progress_callback:
            progress_callback(2, 4, "Reempaquetando DOCX...")

        if cancel_check and cancel_check():
            raise InterruptedError("Operación cancelada")

        # Crear nuevo DOCX
        with zipfile.ZipFile(str(output_docx), 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(str(extract_dir)):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(extract_dir)
                    zipf.write(str(file_path), str(arcname))

    if progress_callback:
        progress_callback(4, 4, "Compresión completada")

    new_size = output_docx.stat().st_size
    reduction = ((original_size - new_size) / original_size) * 100 if original_size > 0 else 0

    return {
        "original_size": original_size,
        "new_size": new_size,
        "reduction_percent": max(0, reduction),
        "images_processed": images_processed
    }


# ===========================================================================
# Conversión de Imágenes
# ===========================================================================

def convert_image(
    input_path: Path,
    output_path: Path,
    output_format: str,
    quality: int = 95,
    resize: Optional[tuple[int, int]] = None,
    maintain_aspect: bool = True,
    overwrite: bool = False
) -> None:
    """Convierte una imagen a otro formato."""
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"El archivo ya existe: {output_path}")

    from PIL import Image

    with Image.open(input_path) as img:
        # Convertir modo si es necesario
        original_mode = img.mode

        # Para ICO necesitamos RGBA
        if output_format.lower() == "ico":
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
        # Para JPEG necesitamos RGB
        elif output_format.lower() in ('jpg', 'jpeg'):
            if img.mode in ('RGBA', 'P', 'LA'):
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if 'A' in img.mode else None)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')

        # Redimensionar si se especificó
        if resize:
            target_w, target_h = resize
            orig_w, orig_h = img.size

            if maintain_aspect:
                ratio_w = target_w / orig_w if target_w > 0 else float('inf')
                ratio_h = target_h / orig_h if target_h > 0 else float('inf')
                ratio = min(ratio_w, ratio_h)

                if ratio < 1:  # Solo reducir, no ampliar
                    new_w = int(orig_w * ratio)
                    new_h = int(orig_h * ratio)
                    img = img.resize((new_w, new_h), Image.LANCZOS)
            else:
                if target_w > 0 and target_h > 0:
                    img = img.resize((target_w, target_h), Image.LANCZOS)

        # Guardar
        output_path.parent.mkdir(parents=True, exist_ok=True)

        save_kwargs: dict[str, Any] = {}

        if output_format.lower() in ('jpg', 'jpeg'):
            save_kwargs['quality'] = quality
            save_kwargs['optimize'] = True
        elif output_format.lower() == 'png':
            save_kwargs['optimize'] = True
        elif output_format.lower() == 'webp':
            save_kwargs['quality'] = quality
        elif output_format.lower() == 'ico':
            # ICO tiene tamaños específicos
            sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
            img.save(str(output_path), format='ICO', sizes=sizes)
            return

        img.save(str(output_path), format=output_format.upper(), **save_kwargs)


def get_image_info(input_path: Path) -> dict:
    """Obtiene información de una imagen."""
    from PIL import Image

    file_size = input_path.stat().st_size

    with Image.open(input_path) as img:
        width, height = img.size
        format_name = img.format or "Unknown"
        mode = img.mode

    # Formatear tamaño
    if file_size < 1024:
        size_str = f"{file_size} B"
    elif file_size < 1024 * 1024:
        size_str = f"{file_size / 1024:.1f} KB"
    else:
        size_str = f"{file_size / (1024 * 1024):.1f} MB"

    return {
        "width": width,
        "height": height,
        "format": format_name,
        "mode": mode,
        "file_size": file_size,
        "file_size_human": size_str
    }


def batch_convert_images(
    input_files: list[Path],
    output_dir: Path,
    output_format: str,
    quality: int = 95,
    resize: Optional[tuple[int, int]] = None,
    maintain_aspect: bool = True,
    overwrite: bool = False,
    progress_callback: Optional[ProgressCallback] = None,
    cancel_check: Optional[CancelCheck] = None
) -> dict:
    """Convierte múltiples imágenes en lote."""
    output_dir.mkdir(parents=True, exist_ok=True)

    total = len(input_files)
    converted = 0
    errors = []

    for i, input_path in enumerate(input_files):
        if cancel_check and cancel_check():
            raise InterruptedError("Operación cancelada")

        if progress_callback:
            progress_callback(i + 1, total, f"Convirtiendo {input_path.name}")

        try:
            ext = output_format.lower()
            if ext == "jpeg":
                ext = "jpg"
            output_path = output_dir / f"{input_path.stem}.{ext}"

            convert_image(
                input_path, output_path, output_format,
                quality, resize, maintain_aspect, overwrite
            )
            converted += 1

        except Exception as e:
            errors.append((input_path.name, str(e)))

    return {
        "total": total,
        "converted": converted,
        "errors": errors
    }


# ===========================================================================
# Extracción de imágenes
# ===========================================================================

def extract_images_from_pdf(
    input_pdf: Path,
    output_dir: Path,
    output_format: str = "png"
) -> list[Path]:
    """Extrae todas las imágenes de un PDF."""
    import fitz
    from PIL import Image
    from io import BytesIO

    output_dir.mkdir(parents=True, exist_ok=True)
    extracted = []

    doc = fitz.open(str(input_pdf))

    try:
        img_count = 0
        for page_num in range(doc.page_count):
            page = doc[page_num]
            image_list = page.get_images()

            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]

                # Convertir a formato deseado
                pil_img = Image.open(BytesIO(image_bytes))

                if output_format.lower() in ('jpg', 'jpeg') and pil_img.mode in ('RGBA', 'P'):
                    pil_img = pil_img.convert('RGB')

                img_count += 1
                ext = "jpg" if output_format.lower() == "jpeg" else output_format.lower()
                output_path = output_dir / f"imagen_{img_count:04d}.{ext}"

                pil_img.save(str(output_path))
                extracted.append(output_path)

    finally:
        doc.close()

    return extracted


def extract_images_from_docx(
    input_docx: Path,
    output_dir: Path
) -> list[Path]:
    """Extrae todas las imágenes de un archivo DOCX."""
    output_dir.mkdir(parents=True, exist_ok=True)
    extracted = []

    with zipfile.ZipFile(str(input_docx), 'r') as zipf:
        for name in zipf.namelist():
            if name.startswith("word/media/"):
                # Extraer imagen
                img_data = zipf.read(name)
                img_name = Path(name).name
                output_path = output_dir / img_name

                output_path.write_bytes(img_data)
                extracted.append(output_path)

    return extracted
