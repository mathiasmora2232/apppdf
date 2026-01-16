# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file para PDF Converter Pro
"""

import sys
from pathlib import Path

block_cipher = None

# Ruta base del proyecto
BASE_PATH = Path(SPECPATH)

# Archivos principales
main_script = str(BASE_PATH / 'convertidor.py')

# Datos adicionales (si hubiera recursos)
datas = []

# Hidden imports necesarios para las librerías
hidden_imports = [
    'pdf2docx',
    'docx',
    'docx2pdf',
    'PIL',
    'PIL.Image',
    'pikepdf',
    'pytesseract',
    'fitz',
    'customtkinter',
    'tkinter',
    'tkinter.filedialog',
    'tkinter.messagebox',
]

a = Analysis(
    [main_script],
    pathex=[str(BASE_PATH)],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PDF Converter Pro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False = sin ventana de consola (GUI app)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Puedes agregar un .ico aquí: icon='icon.ico'
)
