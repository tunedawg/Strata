"""
Strata — launcher (pywebview edition)
Serves templates via a local HTTP server to avoid Edge file:// restrictions
in MSIX sandbox. No Flask — uses Python's built-in http.server.
"""

import os
import sys
import threading
import traceback
import socket
import http.server


# ── 1. Locate bundled resources ───────────────────────────────────────────────
def resource_path(*parts):
    candidates = []
    if hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, *parts))
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    candidates.append(os.path.join(exe_dir, *parts))
    candidates.append(os.path.join(os.path.dirname(exe_dir), *parts))
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.join(script_dir, *parts))
    for path in candidates:
        if os.path.exists(path):
            return path
    return candidates[0]


# ── 2. Find a free port ───────────────────────────────────────────────────────
def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


# ── 3. Start local HTTP server for templates ──────────────────────────────────
def start_template_server(template_dir, port):
    """
    Serve the templates folder on localhost so Edge WebView2
    loads them over http:// instead of file:// — avoids sandbox restrictions.
    """
    os.chdir(template_dir)

    handler = http.server.SimpleHTTPRequestHandler
    # Silence request logs
    handler.log_message = lambda *args: None

    server = http.server.HTTPServer(('127.0.0.1', port), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server


# ── 4. Persistent data directory ─────────────────────────────────────────────
def get_data_dir():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    if not os.path.isdir(docs):
        docs = os.path.expanduser("~")
    data_dir = os.path.join(docs, "Strata")
    try:
        os.makedirs(data_dir, exist_ok=True)
        test = os.path.join(data_dir, ".write_test")
        open(test, "w").close()
        os.remove(test)
    except Exception:
        appdata = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
        data_dir = os.path.join(appdata, "Strata")
        os.makedirs(data_dir, exist_ok=True)
    return data_dir


# ── 5. Set env vars before importing app ─────────────────────────────────────
os.environ["UNIVERSAL_SEARCH_DATA"] = get_data_dir()

if hasattr(sys, "_MEIPASS"):
    sys.path.insert(0, sys._MEIPASS)


# ── 6. Import the API ─────────────────────────────────────────────────────────
try:
    import app as _app_module
    api = _app_module.Api()
except Exception as e:
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Strata — Startup Error",
        f"Failed to load application:\n\n{e}\n\n{traceback.format_exc()}")
    sys.exit(1)


# ── 7. Main ───────────────────────────────────────────────────────────────────
def main():
    import webview

    data_dir     = os.environ["UNIVERSAL_SEARCH_DATA"]
    template_dir = resource_path("templates")

    if not os.path.exists(template_dir):
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Strata — Startup Error",
            f"Could not find application files.\n\nExpected: {template_dir}\n\nPlease reinstall Strata.")
        sys.exit(1)

    # Start local HTTP server so Edge can load the HTML without file:// issues
    port   = find_free_port()
    server = start_template_server(template_dir, port)
    url    = f"http://127.0.0.1:{port}/index.html"

    print("=" * 54)
    print("  Strata")
    print("=" * 54)
    print(f"  Python      : {sys.version.split()[0]}")
    print(f"  Data folder : {data_dir}")
    print(f"  Templates   : {template_dir}")
    print(f"  Server URL  : {url}")
    print("=" * 54)

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
