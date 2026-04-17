# -*- mode: python ; coding: utf-8 -*-
#
# Universal Search — PyInstaller spec
#
# Build with:
#   Windows:  build.bat
#   macOS:    bash build.sh
#
# ── Optional OCR bundling ─────────────────────────────────────────────────────
# To bundle Tesseract + Poppler so end users don't need to install anything:
#
#   Windows:
#     1. Download Tesseract installer from:
#        https://github.com/UB-Mannheim/tesseract/wiki
#        Run it, install to default location, then copy the installed folder to:
#        .\tesseract\   (so tesseract\tesseract.exe exists next to this spec file)
#
#     2. Download Poppler from:
#        https://github.com/oschwartz10612/poppler-windows/releases
#        Extract so that:
#        .\poppler\bin\pdftoppm.exe  exists next to this spec file
#
#   macOS:
#     brew install tesseract poppler
#     The build script will find them automatically via PATH.
#
# If these folders are absent, OCR is still available to users who have
# Tesseract/Poppler installed on their own machine.
# ─────────────────────────────────────────────────────────────────────────────

import os
import sys
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

HERE = os.path.dirname(os.path.abspath(SPEC))

# ── Bundle Tesseract if present locally ──────────────────────────────────────
tesseract_dir = os.path.join(HERE, "tesseract")
bundle_tesseract = os.path.isdir(tesseract_dir) and os.path.isfile(
    os.path.join(tesseract_dir, "tesseract.exe")
)
if bundle_tesseract:
    print(f"[spec] Bundling Tesseract from: {tesseract_dir}")
else:
    print("[spec] No local tesseract/ folder — OCR requires user install")

# ── Bundle Poppler if present locally ────────────────────────────────────────
poppler_bin = os.path.join(HERE, "poppler", "bin")
bundle_poppler = os.path.isdir(poppler_bin) and os.path.isfile(
    os.path.join(poppler_bin, "pdftoppm.exe")
)
if bundle_poppler:
    print(f"[spec] Bundling Poppler from: {poppler_bin}")
else:
    print("[spec] No local poppler/bin/ folder — PDF-to-image OCR unavailable without install")

# ── Collect data files ────────────────────────────────────────────────────────
datas = [
    ("templates", "templates"),
]

if bundle_tesseract:
    # Bundle the entire tesseract folder (includes tessdata language files)
    datas.append((tesseract_dir, "tesseract"))

if bundle_poppler:
    # Bundle all poppler binaries
    datas.append((poppler_bin, "poppler/bin"))

# ── Hidden imports ────────────────────────────────────────────────────────────
hidden_imports = [
    # Flask internals
    "flask", "flask.templating", "jinja2", "jinja2.ext",
    "werkzeug", "werkzeug.routing", "werkzeug.serving",
    "werkzeug.middleware.shared_data",
    "click",
    # PDF
    "pdfplumber", "pdfminer", "pdfminer.high_level", "pdfminer.layout",
    "pdfminer.converter", "pdfminer.pdfinterp", "pdfminer.pdfdevice",
    "pypdf", "pypdf.generic",
    # OCR (optional — app detects Tesseract at runtime)
    "pytesseract", "PIL", "PIL.Image", "PIL.ImageDraw", "pdf2image",
    # Office formats
    "docx", "docx.oxml", "docx.oxml.ns",
    "pptx", "pptx.util",
    "openpyxl", "openpyxl.styles", "openpyxl.utils",
    "xlrd",
    # Email
    "email", "email.utils", "email.header",
    "extract_msg",
    # Report generation
    "reportlab", "reportlab.lib", "reportlab.lib.pagesizes",
    "reportlab.platypus", "reportlab.lib.styles",
    # GUI (splash screen + error dialogs)
    "tkinter", "tkinter.messagebox",
    # Tray icon (Windows — optional, degrades gracefully if absent)
    "pystray",
    # Standard library
    "pickle", "threading", "pathlib", "zipfile", "io", "shutil",
    "xml.etree.ElementTree",
    "charset_normalizer",
    "cryptography",
]

for pkg in ("pdfminer", "pdfplumber", "reportlab", "openpyxl", "jinja2", "werkzeug"):
    hidden_imports += collect_submodules(pkg)

# ── Extra binaries / data from packages ──────────────────────────────────────
extra_datas    = []
extra_binaries = []

for pkg in ("pdfplumber", "pdfminer", "reportlab"):
    d, b, h = collect_all(pkg)
    extra_datas    += d
    extra_binaries += b
    hidden_imports += h

datas += extra_datas

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=extra_binaries,
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib", "numpy", "pandas", "scipy",
        "IPython", "jupyter",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── Executable ────────────────────────────────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="UniversalSearch",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    # console=False → no black terminal window for end users.
    # Set True temporarily if you need to debug a build.
    console=False,
    icon=None,    # replace with "icon.ico" / "icon.icns" for a custom icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="UniversalSearch",
)

# ── macOS .app bundle ─────────────────────────────────────────────────────────
app = BUNDLE(
    coll,
    name="UniversalSearch.app",
    icon=None,
    bundle_identifier="com.universalsearch.app",
    info_plist={
        "CFBundleShortVersionString": "1.0.0",
        "CFBundleVersion":            "1.0.0",
        "NSHighResolutionCapable":    True,
        "LSUIElement":                False,
    },
)
