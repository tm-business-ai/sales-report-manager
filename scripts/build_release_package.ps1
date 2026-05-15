param(
    [string]$AppName = "SalesReportManager",
    [string]$PackageName = "SalesReportManager"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$DistDir = Join-Path $ProjectRoot "dist\$AppName"
$ExePath = Join-Path $DistDir "$AppName.exe"
$ReleaseRoot = Join-Path $ProjectRoot "release"
$PackageDir = Join-Path $ReleaseRoot $PackageName

Set-Location $ProjectRoot

if (-not (Test-Path $DistDir)) {
    throw "EXE directory was not found. Run '.\scripts\build_exe.ps1' first: $DistDir"
}

if (-not (Test-Path $ExePath)) {
    throw "EXE was not found. Run '.\scripts\build_exe.ps1' first: $ExePath"
}

$ResolvedReleaseRoot = [System.IO.Path]::GetFullPath($ReleaseRoot)
$ResolvedPackageDir = [System.IO.Path]::GetFullPath($PackageDir)
if (-not $ResolvedPackageDir.StartsWith($ResolvedReleaseRoot)) {
    throw "Invalid package directory path: $ResolvedPackageDir"
}

if (Test-Path $PackageDir) {
    Remove-Item -LiteralPath $PackageDir -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $PackageDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $PackageDir "data\input") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $PackageDir "data\output") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $PackageDir "docs") | Out-Null

Copy-Item -Path (Join-Path $DistDir "*") -Destination $PackageDir -Recurse -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "docs\README_QUICK_START.txt") -Destination (Join-Path $PackageDir "README_QUICK_START.txt") -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "config.example.json") -Destination (Join-Path $PackageDir "config.example.json") -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "docs\operation_manual.md") -Destination (Join-Path $PackageDir "docs\operation_manual.md") -Force

$SampleFiles = @(
    "sales_2026_03.csv",
    "sales_2026_04.csv",
    "sample_sales_2026_03.xlsx",
    "sample_sales_2026_04.xlsx"
)

foreach ($FileName in $SampleFiles) {
    $Source = Join-Path $ProjectRoot "data\input\$FileName"
    if (Test-Path $Source) {
        Copy-Item -LiteralPath $Source -Destination (Join-Path $PackageDir "data\input\$FileName") -Force
    }
    else {
        Write-Warning "Sample data was not found: $Source"
    }
}

Write-Output "Created release package directory: $PackageDir"
