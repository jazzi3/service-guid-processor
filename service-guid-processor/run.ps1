$ErrorActionPreference = "Stop"

Set-Location $PSScriptRoot

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    Write-Host "Creating virtual environment..."
    python -m venv .venv
}

Write-Host "Installing/updating dependencies..."
& $python -m pip install -r requirements.txt

Start-Process "http://127.0.0.1:8000"
Write-Host "Starting Service GUID Processor at http://127.0.0.1:8000"
& $python -m uvicorn app:app --host 127.0.0.1 --port 8000
