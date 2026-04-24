param(
  [switch]$IncludeDev
)

$ErrorActionPreference = "Stop"

Set-Location (Split-Path $PSScriptRoot -Parent)

if (!(Test-Path ".\venv\Scripts\python.exe")) {
  Write-Host "Creando entorno virtual en .\venv ..."
  python -m venv venv
}

. .\venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
pip install -r requirements.txt

if ($IncludeDev) {
  pip install -r requirements-dev.txt
}

Write-Host "Entorno virtual listo en .\venv"
