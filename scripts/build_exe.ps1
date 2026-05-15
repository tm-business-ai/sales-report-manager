param(
    [string]$AppName = "SalesReportManager",
    [switch]$InstallPyInstaller
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Python = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

Set-Location $ProjectRoot

if (-not (Test-Path $Python)) {
    throw "Python was not found in .venv. Run setup.ps1 first: $Python"
}

$PreviousErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "Continue"
& $Python -c "import PyInstaller" > $null 2>&1
$PyInstallerExitCode = $LASTEXITCODE
$ErrorActionPreference = $PreviousErrorActionPreference

if ($PyInstallerExitCode -ne 0) {
    if ($InstallPyInstaller) {
        & $Python -m pip install pyinstaller
    }
    else {
        throw "PyInstaller was not found. Run '.\.venv\Scripts\python.exe -m pip install -r requirements-optional.txt' and try again."
    }
}

$LegacyOneFileExe = Join-Path $ProjectRoot "dist\$AppName.exe"
if (Test-Path $LegacyOneFileExe) {
    Remove-Item -LiteralPath $LegacyOneFileExe -Force
}

& $Python -m PyInstaller --noconfirm --clean --onedir --noconsole --name $AppName gui.py

$ExePath = Join-Path $ProjectRoot "dist\$AppName\$AppName.exe"
if (-not (Test-Path $ExePath)) {
    throw "Failed to create EXE: $ExePath"
}

Write-Output "Created EXE: $ExePath"
