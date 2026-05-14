$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Python = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

Set-Location $ProjectRoot

& $Python -m pip show pyinstaller *> $null
if ($LASTEXITCODE -ne 0) {
    & $Python -m pip install pyinstaller
}

& $Python -m PyInstaller --noconfirm --onefile --name monthly_report_cli main.py
& $Python -m PyInstaller --noconfirm --onefile --windowed --name monthly_report_gui gui.py

Write-Output "Built executables in dist\"
