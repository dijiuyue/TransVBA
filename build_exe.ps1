# Build single-file executable with PyInstaller
# Requires: pip install pyinstaller

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

# Ensure pyinstaller is available
try {
    pyinstaller --version | Out-Null
} catch {
    Write-Error "PyInstaller not found. Run: pip install pyinstaller"
    exit 1
}

# Build
pyinstaller `
    --onefile `
    --windowed `
    --name "TransVBA-Pro" `
    --add-data "tvba_core_*.py;." `
    --hidden-import win32com.client `
    --hidden-import lxml.etree `
    tvba.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build successful: dist\TransVBA-Pro.exe" -ForegroundColor Green
} else {
    Write-Error "Build failed"
    exit 1
}
