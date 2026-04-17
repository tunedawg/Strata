@echo off
setlocal enabledelayedexpansion
REM ============================================================
REM  Universal Search — Windows Build Script
REM  Output: dist\UniversalSearch_Setup.exe
REM
REM  ── FIRST TIME SETUP ──────────────────────────────────────
REM  Run these once (requires Python 3.9+ on PATH):
REM
REM    pip install pyinstaller flask pdfplumber pypdf python-docx
REM    pip install openpyxl reportlab python-pptx extract-msg
REM    pip install pillow pytesseract pdf2image pystray
REM
REM  ── OPTIONAL: BUNDLE OCR (recommended) ────────────────────
REM  Bundling Tesseract + Poppler means end users get OCR
REM  without installing anything extra.
REM
REM  STEP A — Tesseract:
REM    1. Download installer from:
REM       https://github.com/UB-Mannheim/tesseract/wiki
REM       (get the 64-bit .exe, e.g. tesseract-ocr-w64-setup-5.x.x.exe)
REM    2. Run the installer, use default location
REM    3. Copy the installed folder to .\tesseract\ next to this script
REM       Default install path: C:\Program Files\Tesseract-OCR\
REM       So run: xcopy /E /I "C:\Program Files\Tesseract-OCR" tesseract
REM
REM  STEP B — Poppler:
REM    1. Download from:
REM       https://github.com/oschwartz10612/poppler-windows/releases
REM       (get the latest Release-xx.xx.x-0.zip)
REM    2. Extract it — look inside for a folder containing bin\pdftoppm.exe
REM    3. Copy that folder to .\poppler\ next to this script
REM       So .\poppler\bin\pdftoppm.exe should exist
REM
REM  ── INSTALLER ─────────────────────────────────────────────
REM  Inno Setup 6: https://jrsoftware.org/isdl.php
REM ============================================================

echo.
echo  Universal Search ^— Windows Build
echo  ====================================
echo.

REM ── Locate PyInstaller ───────────────────────────────────────────────────────
where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo  PyInstaller not found ^— installing...
    pip install pyinstaller
    if errorlevel 1 ( echo  ERROR: pip install failed. & pause & exit /b 1 )
)

set PI=pyinstaller
where pyinstaller >nul 2>&1
if errorlevel 1 (
    for /f "delims=" %%i in ('python -c "import sys,os; print(os.path.join(os.path.dirname(sys.executable),\"Scripts\",\"pyinstaller.exe\"))"') do set PI=%%i
)

REM ── Report OCR bundling status ───────────────────────────────────────────────
echo  Checking optional OCR components...
if exist tesseract\tesseract.exe (
    echo  [OK] Tesseract found ^— will be bundled ^(users get OCR out of the box^)
) else (
    echo  [--] No tesseract\ folder ^— OCR requires Tesseract installed on user machine
    echo       To bundle: see OPTIONAL steps in this script header
)
if exist poppler\bin\pdftoppm.exe (
    echo  [OK] Poppler found ^— will be bundled
) else (
    echo  [--] No poppler\bin\ folder ^— PDF-to-image OCR unavailable without user install
)
echo.

REM ── Clean previous build ─────────────────────────────────────────────────────
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist

REM ── Run PyInstaller ──────────────────────────────────────────────────────────
echo  Building application bundle...
"%PI%" universal_search.spec --noconfirm

if errorlevel 1 (
    echo.
    echo  ERROR: PyInstaller build failed. See output above.
    pause
    exit /b 1
)

echo.
echo  Bundle created: dist\UniversalSearch\

REM ── Run Inno Setup ───────────────────────────────────────────────────────────
set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"

if exist %ISCC% (
    echo  Building installer...
    %ISCC% installer.iss
    if errorlevel 1 (
        echo  ERROR: Inno Setup failed. Check output above.
    ) else (
        echo.
        echo  ============================================================
        echo   SUCCESS: dist\UniversalSearch_Setup.exe is ready to share.
        echo  ============================================================
    )
) else (
    echo.
    echo  Inno Setup not found ^— skipping installer creation.
    echo  ^> Install from: https://jrsoftware.org/isdl.php
    echo  ^> Or distribute the dist\UniversalSearch\ folder directly ^(zipped^).
)

echo.
pause
