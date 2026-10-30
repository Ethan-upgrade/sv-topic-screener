# setup (Windows / PowerShell) - create venv and install deps
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
Write-Host "=== sv-topic-screener setup ==="
try { Write-Host ("python: " + (& python --version 2>&1)) } catch { Write-Host "ERROR: Python 3.11+ not found in PATH. Install it first."; exit 1 }
if (-not (Test-Path ".venv")) { Write-Host "creating venv .venv ..."; & python -m venv .venv }
$vpy = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
Write-Host "installing dependencies (large, needs internet) ..."
& $vpy -m pip install --upgrade pip
& $vpy -m pip install -r requirements.txt
Write-Host ""
Write-Host "Done. Quick test:"
Write-Host "  .venv\Scripts\python.exe screener.py --init-example"
Write-Host "  .venv\Scripts\python.exe screener.py --input example_input.xlsx"
