param(
    [switch]$OneFile = $true
)

$ErrorActionPreference = 'Stop'

Write-Host "[Build] Creando entorno..."
python -m pip install --upgrade pip | Out-Null
python -m pip install -r requirements.txt | Out-Null
python -m pip install pyinstaller | Out-Null

$argsList = @(
    "--noconfirm",
    "--windowed",
    "--name", "convertidor",
    "--hidden-import", "customtkinter",
    "--collect-all", "customtkinter"
)
if ($OneFile) { $argsList += "--onefile" }

Write-Host "[Build] Ejecutando PyInstaller..."
pyinstaller @argsList "convertidor.py"

Write-Host "[Build] Listo. Ejecutable en ./dist/convertidor.exe"
