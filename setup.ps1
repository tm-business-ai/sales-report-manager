param(
    [switch]$InstallDragDrop
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$PythonCommand = "python"

Set-Location $ProjectRoot

if (-not (Test-Path ".venv")) {
    & $PythonCommand -m venv .venv
}

& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r requirements.txt

if ($InstallDragDrop) {
    & $VenvPython -m pip install tkinterdnd2
}

New-Item -ItemType Directory -Force -Path "data\input" | Out-Null
New-Item -ItemType Directory -Force -Path "data\output" | Out-Null
New-Item -ItemType Directory -Force -Path "logs" | Out-Null

& $VenvPython main.py --check-setup
