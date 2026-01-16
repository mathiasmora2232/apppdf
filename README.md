# Convertidor PDF a Word (DOCX)

Herramienta sencilla en Python para convertir archivos PDF a documentos Word (DOCX), con soporte para conversión por lotes y rango de páginas.

## Requisitos
- Python 3.9+
- Windows (probado); también funciona en macOS/Linux
- Paquetes: `pdf2docx`, `python-docx` (se instalan más abajo)

## Instalación

```powershell
# Desde la carpeta del proyecto
python -m pip install -r requirements.txt
```

## Modo GUI (interfaz gráfica)

Ejecuta la interfaz para seleccionar un PDF y convertirlo sin usar la terminal.

```powershell
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe convertidor.py
```

En la ventana podrás:
- PDF → DOCX: elegir PDF, salida DOCX opcional, rango de páginas y sobrescribir.
- Fidelidad exacta (imagen): genera DOCX idéntico visualmente al PDF (menos editable).
- DOCX → PDF: elegir DOCX y salida PDF (usa Microsoft Word vía docx2pdf si está disponible).
- Compresión: optimizar PDF (limpieza y deflate) y comprimir imágenes en DOCX (calidad y tamaño máximo).
- Lotes: agregar múltiples archivos (PDF/DOCX), elegir carpeta de salida y convertir en bloque.

## Uso rápido

### Convertir un solo PDF
```powershell
python main.py "ruta/al/archivo.pdf" -o "ruta/salida.docx"
```
Si no especificas `-o`, se generará un `.docx` junto al PDF.

### Convertir todos los PDF de una carpeta
```powershell
python main.py "ruta/a/carpeta" --outdir "ruta/salida"
```

### Convertir un rango de páginas (1-basado)
```powershell
python main.py "archivo.pdf" -o "salida.docx" --start 2 --end 5
```

### Sobrescribir si el DOCX existe
```powershell
python main.py "archivo.pdf" -o "salida.docx" --overwrite
```

## CLI avanzado (múltiples funciones)

```powershell
# Ayuda general
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe cli.py --help

# PDF → DOCX
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe cli.py pdf2docx "input.pdf" -o "output.docx" --start 1 --end 3 --overwrite

# DOCX → PDF (requiere Microsoft Word instalado en Windows)
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe cli.py docx2pdf "input.docx" -o "output.pdf" --overwrite

# Optimizar PDF (reduce tamaño limpiando y deflating)
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe cli.py compress-pdf "input.pdf" -o "optimized.pdf"

# Comprimir imágenes dentro de DOCX
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe cli.py compress-docx "input.docx" -o "compressed.docx" --quality 70 --max-width 1600 --max-height 1200

# Procesar por lotes en una carpeta (PDF/DOCX)
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe cli.py batch "C:\ruta\carpeta" --outdir "C:\ruta\salida" --pdf2docx --docx2pdf --overwrite
# Modo fidelidad exacta (imagen) para PDFs
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe cli.py batch "C:\ruta\carpeta" --outdir "C:\ruta\salida" --pdf2docx-raster --dpi 200
```

## Limitaciones y notas

## OCR (PDF imagen → DOCX texto)

Requisitos adicionales:
- Instalar Tesseract OCR:
	- Winget: `winget install -e --id UB-Mannheim.TesseractOCR`
	- O descarga: https://github.com/tesseract-ocr/tesseract
- Ajustar idioma: por defecto `eng`; para español usa `spa` si tienes el paquete instalado.
	- Si ves el error de `spa.traineddata`, instala el idioma español o usa `--lang eng`.

CLI:
```powershell
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe cli.py ocr-pdf2docx "input.pdf" -o "output.docx" --dpi 300 --lang eng
# Idioma mixto (si instalados): --lang spa+eng
```

GUI:
- En la pestaña PDF → DOCX, usa “Convertir (OCR texto)”, define `OCR idioma` y `OCR DPI`.
## Estructura
- `main.py`: CLI del convertidor
- `gui.py`: Interfaz gráfica con Tkinter
- `requirements.txt`: dependencias
- `README.md`: instrucciones

## Ejemplos adicionales
- Carpeta sin `--outdir`: los `.docx` se crean en la misma carpeta.
- Rango parcial: si `--end` se omite, se convierte desde `--start` hasta el final.

## Problemas comunes
- "FileExistsError": usa `--overwrite` o cambia la ruta de salida.
- "FileNotFoundError": verifica la ruta del PDF o carpeta.
- Si ves errores de importación, reinstala dependencias:
```powershell
python -m pip install --upgrade pdf2docx python-docx
```
