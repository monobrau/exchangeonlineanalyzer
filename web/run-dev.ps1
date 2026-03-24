#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here
if (-not (Test-Path '.venv\Scripts\python.exe')) {
    Write-Host 'Run: python -m venv .venv; .venv\Scripts\pip install -r requirements.txt' -ForegroundColor Yellow
    exit 1
}
& .\.venv\Scripts\python.exe -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8080
