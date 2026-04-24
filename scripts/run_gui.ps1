$ErrorActionPreference = "Stop"

Set-Location (Split-Path $PSScriptRoot -Parent)

if (!(Test-Path ".\venv\Scripts\python.exe")) {
  & .\scripts\setup_venv.ps1
}
else {
  . .\venv\Scripts\Activate.ps1
}

python .\assemble_gui.py
