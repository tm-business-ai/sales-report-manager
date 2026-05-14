$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Python = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$Config = Join-Path $ProjectRoot "config.example.json"
$Log = Join-Path $ProjectRoot "logs\scheduled_task.log"

Set-Location $ProjectRoot

& $Python main.py --config $Config --preset april_report --notify *> $Log
exit $LASTEXITCODE
