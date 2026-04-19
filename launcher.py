"""
Strata — launcher (pywebview edition)
Opens a native desktop window — no browser, no Flask, no port needed.
"""

import os
import sys
import threading
import traceback

# ── 1. Locate bundled resources ───────────────────────────────────────────────
def resource_path(*parts):
    """
    Find a resource file, handling:
    - Running from source (dev)
    - PyInstaller bundle (_MEIPASS)
    - MSIX sandbox (WindowsApps, VFS)
    """
    candidates = []

    # PyInstaller bundle
    if hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, *parts))

    # Exe directory and its parent
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    candidates.append(os.path.join(exe_dir, *parts))
    candidates.append(os.path.join(os.path.dirname(exe_dir), *parts))

    # Script directory (dev mode)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.join(script_dir, *parts))

    # MSIX VFS path — Windows virtualises the install dir here
    localappdata = os.environ.get("LOCALAPPDATA", "")
    if localappdata:
        candidates.append(os.path.join(
            localappdata, "Microsoft", "WindowsApps", *parts))

    for path in candidates:
        if os.path.exists(path):
            return path

    return candidates[0]  # best guess


# ── 2. Persistent data directory ─────────────────────────────────────────────
def get_data_dir():
    """
    Returns a writable data directory. MSIX sandbox restricts writes to
    AppData and Documents — we use Documents\Strata which is always allowed.
    """
    # Try standard Documents folder
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    if not os.path.isdir(docs):
        docs = os.path.expanduser("~")

    data_dir = os.path.join(docs, "Strata")
    try:
        os.makedirs(data_dir, exist_ok=True)
        # Verify it's actually writable
        test = os.path.join(data_dir, ".write_test")
        with open(test, "w") as f:
            f.write("ok")
        os.remove(test)
    except Exception:
        # Fallback to AppData\Local\Strata (also allowed in MSIX)
        appdata = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
        data_dir = os.path.join(appdata, "Strata")
        os.makedirs(data_dir, exist_ok=True)

    return data_dir


# ── 3. Set env vars before importing app ─────────────────────────────────────
os.environ["UNIVERSAL_SEARCH_DATA"] = get_data_dir()

if hasattr(sys, "_MEIPASS"):
    sys.path.insert(0, sys._MEIPASS)


# ── 4. Import the API ─────────────────────────────────────────────────────────
try:
    import app as _app_module
    api = _app_module.Api()
except Exception as e:
    # Show error dialog if API fails to load
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Strata — Startup Error",
        f"Failed to load application:\n\n{e}\n\n{traceback.format_exc()}")
    sys.exit(1)


# ── 5. Find index.html ────────────────────────────────────────────────────────
def find_index_html():
    # Try resource_path first
    path = resource_path("templates", "index.html")
    if os.path.exists(path):
        return path

    # Walk from exe dir looking for index.html
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    for root, dirs, files in os.walk(exe_dir):
        # Don't recurse into AppData or Windows system folders
        dirs[:] = [d for d in dirs if d not in
                   ("AppData", "Windows", "System32", "SysWOW64")]
        if "index.html" in files:
            return os.path.join(root, "index.html")

    return None


# ── 6. Main ───────────────────────────────────────────────────────────────────
def main():
    import webview

    data_dir   = os.environ["UNIVERSAL_SEARCH_DATA"]
    index_html = find_index_html()

    print("=" * 54)
    print("  Strata")
    print("=" * 54)
    print(f"  Python      : {sys.version.split()[0]}")
    print(f"  Executable  : {sys.executable}")
    print(f"  Data folder : {data_dir}")
    print(f"  Index HTML  : {index_html}")
    print(f"  HTML exists : {os.path.exists(index_html) if index_html else False}")
    print("=" * 54)

    if not index_html or not os.path.exists(index_html):
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Strata — Startup Error",
            "Could not find application files.\n\n"
            "Please reinstall Strata.")
        sys.exit(1)

    # Normalise path for file:/// URL (forward slashes, no leading slash issues)
    url = "file:///" + index_html.replace("\\", "/").lstrip("/")

    window = webview.create_window(
        title     = "Strata",
        url       = url,
        js_api    = api,
        width     = 1280,
        height    = 860,
        min_size  = (900, 600),
        resizable = True,
    )

    webview.start(debug=False)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Strata — Error",
            f"Unexpected error:\n\n{e}\n\n{traceback.format_exc()}")
