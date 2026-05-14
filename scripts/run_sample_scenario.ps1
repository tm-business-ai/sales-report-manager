$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Python = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$OutputName = "sample_scenario_report.xlsx"

Set-Location $ProjectRoot

& $Python main.py --check-setup
& $Python main.py --config config.example.json --preset april_report --preview --preview-limit 3
& $Python main.py --config config.example.json --preset april_report --output-name $OutputName --summary-csv-prefix sample_scenario

Write-Output "Sample scenario completed. Output: data\output\$OutputName"
