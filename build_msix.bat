@echo off
REM ============================================================
REM  Strata — MSIX Package Builder
REM  Run AFTER build.bat has produced dist\Strata\
REM
REM  Requires Windows SDK (makeappx.exe + signtool.exe)
REM  Install from: https://developer.microsoft.com/windows/downloads/windows-sdk/
REM ============================================================

echo.
echo  Strata ^— MSIX Builder
echo  ====================================
echo.

REM ── Find makeappx.exe ────────────────────────────────────────────────────────
set SDKROOT=C:\Program Files (x86)\Windows Kits\10\bin
set MAKEAPPX=
for /d %%v in ("%SDKROOT%\10.*") do (
    if exist "%%v\x64\makeappx.exe" set MAKEAPPX=%%v\x64\makeappx.exe
)
if "%MAKEAPPX%"=="" (
    echo  ERROR: makeappx.exe not found. Install Windows SDK.
    echo  https://developer.microsoft.com/windows/downloads/windows-sdk/
    pause & exit /b 1
)
echo  Found: %MAKEAPPX%

REM ── Find signtool.exe ────────────────────────────────────────────────────────
set SIGNTOOL=
for /d %%v in ("%SDKROOT%\10.*") do (
    if exist "%%v\x64\signtool.exe" set SIGNTOOL=%%v\x64\signtool.exe
)

REM ── Build MSIX staging folder ────────────────────────────────────────────────
echo  Preparing MSIX staging folder...
if exist msix_staging rmdir /s /q msix_staging
mkdir msix_staging

REM Copy app files
xcopy /E /I /Q dist\Strata msix_staging
REM Copy manifest and assets
copy AppxManifest.xml msix_staging\AppxManifest.xml
xcopy /E /I /Q assets msix_staging\assets

REM ── Create MSIX ──────────────────────────────────────────────────────────────
echo  Creating MSIX package...
if exist dist\Strata.msix del dist\Strata.msix

"%MAKEAPPX%" pack /d msix_staging /p dist\Strata.msix /nv
if errorlevel 1 (
    echo  ERROR: makeappx failed.
    pause & exit /b 1
)
echo  Created: dist\Strata.msix

REM ── Sign with self-signed cert (for local testing) ───────────────────────────
if not exist strata_test.pfx (
    echo.
    echo  Creating self-signed certificate for testing...
    powershell -Command "New-SelfSignedCertificate -Type Custom -Subject 'CN=StrataTest' -KeyUsage DigitalSignature -FriendlyName 'Strata Test Cert' -CertStoreLocation 'Cert:\CurrentUser\My' -TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.3','2.5.29.19={text}') | Export-PfxCertificate -FilePath strata_test.pfx -Password (ConvertTo-SecureString -String 'StrataTest123' -Force -AsPlainText)"
)

if exist "%SIGNTOOL%" (
    echo  Signing MSIX...
    "%SIGNTOOL%" sign /fd SHA256 /a /f strata_test.pfx /p StrataTest123 dist\Strata.msix
    if errorlevel 1 (
        echo  WARNING: Signing failed ^— MSIX created but unsigned.
    ) else (
        echo  Signed successfully.
    )
)

echo.
echo  ============================================================
echo   MSIX ready: dist\Strata.msix
echo.
echo   To install locally for testing:
echo     1. Double-click dist\Strata.msix
echo     2. Or: Add-AppxPackage dist\Strata.msix  ^(PowerShell^)
echo.
echo   For Store submission: upload dist\Strata.msix to
echo   Partner Center ^(Microsoft signs it for you^)
echo  ============================================================
echo.
pause
