$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "Virtual environment not found. Run: python -m venv .venv; .\.venv\Scripts\pip install -e .[dev]"
    exit 1
}
& ".venv\Scripts\python.exe" "scheduler.py"
