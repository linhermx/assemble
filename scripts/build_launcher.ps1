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

if (Test-Path ".\dist\launcher") { Remove-Item ".\dist\launcher" -Recurse -Force }

if (!(Test-Path $iconPath)) {
  throw "No se encontro el icono: $iconPath"
}

pyinstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --icon $iconPath `
  --name assemble_launcher `
  assemble_launcher.py

Write-Host "Launcher EXE generado en: dist\assemble_launcher.exe"
