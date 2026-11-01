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

Write-Host ""
Write-Host "[env] optional JS runtime for yt-dlp (needed for YouTube) ..."
$deno = Get-Command deno -ErrorAction SilentlyContinue
$node = Get-Command node -ErrorAction SilentlyContinue
if ($deno) { Write-Host ("  deno: " + ((& deno --version 2>&1) | Select-Object -First 1)); Write-Host "  use: --js-runtime deno" }
elseif ($node) { Write-Host ("  node: " + (& node --version 2>&1)); Write-Host "  use: --js-runtime node" }
else { Write-Host "  WARN: no deno/node found. For YouTube install Deno (winget install DenoLand.Deno) or Node.js, then use --js-runtime." }
$ff = Get-Command ffmpeg -ErrorAction SilentlyContinue
if (-not $ff) { Write-Host "  WARN: ffmpeg not found in PATH (recommended for some formats)." }
Write-Host ""
