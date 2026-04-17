"""
Universal Search — launcher
Wraps the Flask app for standalone desktop use.

When run as a PyInstaller bundle:
  - Sets data directory to ~/Documents/UniversalSearch (persists across updates)
  - Shows a brief "Starting…" splash so users know something is happening
  - Waits for Flask to be ready, then opens the default browser
  - A system tray icon (Windows) keeps it running; quit via tray or Ctrl+C

When run as plain Python (development):
  - Behaves exactly like: python app.py
"""

import os
import sys
import socket
import threading
import time
import webbrowser

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


# ── 3. Set environment before importing app ───────────────────────────────────
os.environ["UNIVERSAL_SEARCH_DATA"] = get_data_dir()

if hasattr(sys, "_MEIPASS"):
    os.environ["UNIVERSAL_SEARCH_TEMPLATES"] = resource_path("templates")
    sys.path.insert(0, sys._MEIPASS)


# ── 4. Import the Flask application ──────────────────────────────────────────
import app as _app_module

flask_app = _app_module.app

if hasattr(sys, "_MEIPASS"):
    flask_app.template_folder = resource_path("templates")


# ── 5. Port / readiness helpers ───────────────────────────────────────────────
PORT = 5000


def _is_port_free(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) != 0


def _wait_for_flask(port, timeout=20):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.15)
    return False


# ── 6. GUI helpers ────────────────────────────────────────────────────────────
def _show_splash():
    """Tiny 'Starting…' window so users know the app is launching."""
    try:
        import tkinter as tk
        root = tk.Tk()
        root.title("Universal Search")
        root.resizable(False, False)
        w, h = 320, 90
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        root.configure(bg="#0d0f14")
        try:
            root.overrideredirect(True)   # borderless on Windows
        except Exception:
            pass
        tk.Label(root, text="⚙  Universal Search",
                 font=("Segoe UI", 13, "bold"),
                 fg="#4f8ef7", bg="#0d0f14").pack(pady=(18, 4))
        tk.Label(root, text="Starting, please wait…",
                 font=("Segoe UI", 10),
                 fg="#8892a4", bg="#0d0f14").pack()
        root.update()
        return root
    except Exception:
        return None


def _close_splash(root):
    try:
        if root:
            root.destroy()
    except Exception:
        pass


def _show_error(title, message):
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message)
        root.destroy()
    except Exception:
        print(f"ERROR — {title}: {message}", file=sys.stderr)


# ── 7. System tray (Windows — optional, degrades gracefully) ─────────────────
def _start_tray(stop_event):
    try:
        import pystray
        from PIL import Image, ImageDraw
        img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
        d   = ImageDraw.Draw(img)
        d.ellipse([4, 4, 60, 60], fill="#4f8ef7")

        def on_open(_icon, _item):
            webbrowser.open(f"http://localhost:{PORT}")

        def on_quit(_icon, _item):
            _icon.stop()
            stop_event.set()

        icon = pystray.Icon(
            "UniversalSearch", img, "Universal Search",
            menu=pystray.Menu(
                pystray.MenuItem("Open in Browser", on_open, default=True),
                pystray.MenuItem("Quit", on_quit),
            ),
        )
        icon.run()
    except Exception:
        stop_event.wait()


# ── 8. Main ───────────────────────────────────────────────────────────────────
def main():
    frozen = hasattr(sys, "_MEIPASS")

    # Already running — just open browser
    if not _is_port_free(PORT):
        webbrowser.open(f"http://localhost:{PORT}")
        return

    splash = _show_splash() if frozen else None

    stop_event = threading.Event()

    flask_thread = threading.Thread(
        target=lambda: flask_app.run(
            host="127.0.0.1",
            port=PORT,
            debug=False,
            use_reloader=False,
            threaded=True,
        ),
        daemon=True,
    )
    flask_thread.start()

    ready = _wait_for_flask(PORT, timeout=20)
    _close_splash(splash)

    if not ready:
        _show_error(
            "Universal Search — Startup Error",
            f"The application failed to start on port {PORT}.\n\n"
            "Another program may be using that port.\n"
            "Please close other apps and try again.",
        )
        return

    webbrowser.open(f"http://localhost:{PORT}")

    if frozen:
        if sys.platform == "win32":
            _start_tray(stop_event)
        else:
            # macOS: stay alive until the user quits the process
            try:
                while not stop_event.is_set():
                    time.sleep(1)
            except KeyboardInterrupt:
                pass
    else:
        data_dir = os.environ["UNIVERSAL_SEARCH_DATA"]
        print("=" * 54)
        print("  Universal Search  (dev mode)")
        print("=" * 54)
        print(f"  Data folder : {data_dir}")
        print(f"  Address     : http://localhost:{PORT}")
        print(f"  Stop        : Ctrl+C")
        print("=" * 54)
        try:
            flask_thread.join()
        except KeyboardInterrupt:
            print("\n[Universal Search] Shutting down.")


if __name__ == "__main__":
    main()
