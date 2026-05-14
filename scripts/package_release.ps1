param(
    [string]$PackageName = "monthly_report_tool",
    [switch]$IncludeSource
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$ReleaseRoot = Join-Path $ProjectRoot "release"
$PackageDir = Join-Path $ReleaseRoot $PackageName

Set-Location $ProjectRoot

if (Test-Path $PackageDir) {
    Remove-Item -LiteralPath $PackageDir -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $PackageDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $PackageDir "data\input") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $PackageDir "data\output") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $PackageDir "logs") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $PackageDir "scripts") | Out-Null

$CommonFiles = @(
    "README.md",
    "config.example.json",
    "requirements.txt",
    "requirements-optional.txt",
    "setup.ps1",
    "run_gui.bat",
    "run_report.bat"
)

foreach ($File in $CommonFiles) {
    if (Test-Path $File) {
        Copy-Item -LiteralPath $File -Destination $PackageDir
    }
}

Copy-Item -LiteralPath "scripts\register_monthly_task.ps1" -Destination (Join-Path $PackageDir "scripts")
Copy-Item -LiteralPath "scripts\run_monthly_report.ps1" -Destination (Join-Path $PackageDir "scripts")
Copy-Item -LiteralPath "scripts\run_sample_scenario.ps1" -Destination (Join-Path $PackageDir "scripts")

if (Test-Path "dist") {
    New-Item -ItemType Directory -Force -Path (Join-Path $PackageDir "dist") | Out-Null
    Copy-Item -LiteralPath "dist\*.exe" -Destination (Join-Path $PackageDir "dist") -ErrorAction SilentlyContinue
}

if ($IncludeSource -or -not (Test-Path "dist\monthly_report_cli.exe") -or -not (Test-Path "dist\monthly_report_gui.exe")) {
    Copy-Item -LiteralPath "main.py" -Destination $PackageDir
    Copy-Item -LiteralPath "report.py" -Destination $PackageDir
    Copy-Item -LiteralPath "gui.py" -Destination $PackageDir
}

Copy-Item -LiteralPath "data\input\sample_*.csv" -Destination (Join-Path $PackageDir "data\input") -ErrorAction SilentlyContinue
Copy-Item -LiteralPath "data\input\sales_*.csv" -Destination (Join-Path $PackageDir "data\input") -ErrorAction SilentlyContinue

Compress-Archive -LiteralPath $PackageDir -DestinationPath (Join-Path $ReleaseRoot "$PackageName.zip") -Force

Write-Output "Created release package: release\$PackageName.zip"
