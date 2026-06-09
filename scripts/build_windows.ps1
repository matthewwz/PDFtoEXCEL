# Build a standalone Windows .exe (no Python install needed on target PCs).
# From repo root:  .\scripts\build_windows.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot\..

python -m pip install --upgrade pip
pip install -r requirements.txt -r requirements-build.txt

$piArgs = @(
    "--noconfirm", "--clean",
    "--windowed",
    "--onefile",
    "--collect-all", "pdfplumber",
    "--collect-all", "pypdfium2",
    "--name", "PDFtoExcel",
    "gui.py"
)

python -m PyInstaller @piArgs

Write-Host ""
Write-Host "Done. Executable: .\dist\PDFtoExcel.exe"
Get-ChildItem dist\PDFtoExcel.exe | Format-Table Name, Length
