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
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name "TransVBA-Pro" `
    --add-data "templates;templates" `
    --hidden-import win32com.client `
    --hidden-import win32timezone `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --hidden-import lxml.etree `
    tvba.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build successful: dist\TransVBA-Pro.exe" -ForegroundColor Green
} else {
    Write-Error "Build failed"
    exit 1
}
