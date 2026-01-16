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
C:/Users/USER/Desktop/programs/apppdf/.venv/Scripts/python.exe gui.py
```

En la ventana podrás:
- PDF → DOCX: elegir PDF, salida DOCX opcional, rango de páginas y sobrescribir.
- DOCX → PDF: elegir DOCX y salida PDF (usa Microsoft Word vía docx2pdf si está disponible).
- Compresión: optimizar PDF (limpieza y deflate) y comprimir imágenes en DOCX (calidad y tamaño máximo).

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
```

## Limitaciones y notas
- Este conversor no realiza OCR: los PDF escaneados como imagen no extraerán texto.
- La fidelidad del diseño puede variar según el contenido del PDF.
- Para máxima fidelidad de DOCX → PDF en Windows se usa Microsoft Word (docx2pdf). Si no está instalado, fallará.
- Para OCR, puedes evaluar instalar Tesseract y usar `pytesseract` + `pdf2image` (no incluido por simplicidad).

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
