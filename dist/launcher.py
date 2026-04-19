"""
Universal Search — launcher (pywebview edition)
Opens a native desktop window — no browser, no Flask, no port needed.
"""

import os
import sys
import threading

# ── 1. Locate bundled resources ───────────────────────────────────────────────
def resource_path(*parts):
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)

# ── 2. Persistent data directory ─────────────────────────────────────────────
def get_data_dir():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    if not os.path.isdir(docs):
        docs = os.path.expanduser("~")
    data_dir = os.path.join(docs, "UniversalSearch")
    os.makedirs(data_dir, exist_ok=True)
    return data_dir

# ── 3. Set data dir before importing app ─────────────────────────────────────
os.environ["UNIVERSAL_SEARCH_DATA"] = get_data_dir()

if hasattr(sys, "_MEIPASS"):
    sys.path.insert(0, sys._MEIPASS)

# ── 4. Import the API ─────────────────────────────────────────────────────────
import app as _app_module
api = _app_module.Api()

# ── 5. Main ───────────────────────────────────────────────────────────────────
def main():
    import webview

    data_dir     = os.environ["UNIVERSAL_SEARCH_DATA"]
    template_dir = resource_path("templates")
    index_html   = os.path.join(template_dir, "index.html")

    print("=" * 54)
    print("  Strata")
    print("=" * 54)
    print(f"  Data folder : {data_dir}")
    print(f"  Templates   : {template_dir}")
    print("=" * 54)

    window = webview.create_window(
        title      = "Strata",
        url        = f"file:///{index_html.replace(os.sep, '/')}",
        js_api     = api,
        width      = 1280,
        height     = 860,
        min_size   = (900, 600),
        resizable  = True,
    )

    webview.start(debug=True)


if __name__ == "__main__":
    main()
