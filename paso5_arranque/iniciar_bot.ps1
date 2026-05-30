# Arranca el bot desde la raiz del repo con el Python del venv.
$ErrorActionPreference = "Stop"
$botDir = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $botDir
$py = Join-Path $botDir ".venv\Scripts\python.exe"

Set-Location -LiteralPath $repoRoot

if (-not (Test-Path -LiteralPath $py)) {
    Write-Host "Falta el venv. Ejecuta una vez en:" $botDir
    Write-Host "  python -m venv .venv"
    Write-Host "  .\.venv\Scripts\Activate.ps1"
    Write-Host "  pip install -r requirements.txt"
    exit 1
}

& $py -m bot_financiero_telegram
