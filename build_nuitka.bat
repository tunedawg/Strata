@echo off
REM ============================================================
REM  Strata — Nuitka Build Script
REM  Compiles Python to native C executable
REM  Requires: py -3.12 -m pip install nuitka
REM            Visual Studio Build Tools with C++ workload
REM ============================================================

echo.
echo  Strata ^— Nuitka Build
echo  ====================================
echo.

set PY=py -3.12
set ROOT=%~dp0
if "%ROOT:~-1%"=="\" set ROOT=%ROOT:~0,-1%

REM ── Clean ────────────────────────────────────────────────────────────────────
if exist "%ROOT%\strata.dist"  rmdir /s /q "%ROOT%\strata.dist"
if exist "%ROOT%\strata.build" rmdir /s /q "%ROOT%\strata.build"
if exist "%ROOT%\dist"         rmdir /s /q "%ROOT%\dist"
mkdir "%ROOT%\dist"

REM ── Build with Nuitka ────────────────────────────────────────────────────────
echo  Compiling with Nuitka (this takes several minutes)...
%PY% -m nuitka ^
  --standalone ^
  --onefile ^
  --windows-console-mode=disable ^
  --windows-icon-from-ico=strata.ico ^
  --output-filename=Strata.exe ^
  --output-dir=dist ^
  --include-data-dir=templates=templates ^
  --include-package=webview ^
  --nofollow-import-to=webview.platforms.android ^
  --nofollow-import-to=webview.platforms.cocoa ^
  --nofollow-import-to=webview.platforms.gtk ^
  --include-package=pdfplumber ^
  --include-package=pypdf ^
  --include-package=docx ^
  --include-package=openpyxl ^
  --include-package=reportlab ^
  --include-package=pptx ^
  --include-package=PIL ^
  --include-package=pytesseract ^
  --include-package=pdf2image ^
  --include-package=extract_msg ^
  --include-package=pdfminer ^
  --include-package=cryptography ^
  --include-package=charset_normalizer ^
  --assume-yes-for-downloads ^
  launcher.py

if errorlevel 1 (
    echo.
    echo  ERROR: Nuitka build failed.
    pause & exit /b 1
)

echo.
echo  Nuitka build complete: dist\Strata.exe

REM ── Bundle Tesseract if present ──────────────────────────────────────────────
if exist "%ROOT%\tesseract" (
    echo  Copying Tesseract...
    xcopy /E /I /Q "%ROOT%\tesseract" "%ROOT%\dist\tesseract"
)

REM ── Bundle Poppler if present ────────────────────────────────────────────────
if exist "%ROOT%\poppler" (
    echo  Copying Poppler...
    xcopy /E /I /Q "%ROOT%\poppler" "%ROOT%\dist\poppler"
)

REM ── Run Inno Setup ───────────────────────────────────────────────────────────
set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"

if exist %ISCC% (
    echo  Building installer...
    REM Update installer to point to Nuitka output
    %ISCC% installer_nuitka.iss
    if errorlevel 1 (
        echo  WARNING: Inno Setup failed.
    ) else (
        echo.
        echo  ============================================================
        echo   SUCCESS: dist\Strata_Setup.exe
        echo  ============================================================
    )
) else (
    echo  Inno Setup not found ^— distributable exe at dist\Strata.exe
)

echo.
pause
