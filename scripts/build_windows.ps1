param(
  [switch]$Release
)

$ErrorActionPreference = "Stop"

Set-Location (Split-Path $PSScriptRoot -Parent)

$iconPath = ".\src\assemble\resources\assemble.ico"

if (!(Test-Path ".\venv\Scripts\python.exe")) {
  python -m venv venv
}

. .\venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install -r requirements-dev.txt

if (Test-Path ".\build") { Remove-Item ".\build" -Recurse -Force }
if (Test-Path ".\dist") { Remove-Item ".\dist" -Recurse -Force }

if (!(Test-Path $iconPath)) {
  throw "No se encontro el icono: $iconPath"
}

pyinstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --paths ".\src" `
  --hidden-import assemble `
  --hidden-import assemble.gui `
  --hidden-import assemble.core `
  --add-data "src\assemble\resources\assemble.ico;resources" `
  --icon $iconPath `
  --name assemble_windows `
  assemble_gui.py

if (!(Test-Path ".\dist\assemble_windows.exe")) {
  throw "No se genero el ejecutable: dist\assemble_windows.exe"
}

Write-Host "EXE generado en: dist\assemble_windows.exe"

if ($Release) {
  Copy-Item ".\dist\assemble_windows.exe" -Destination ".\dist\assemble.exe" -Force
  Write-Host "Alias listo: dist\assemble.exe"
}
