"""JJ's Setlist helper - runs the browser edition at http://localhost:8765/
and keeps its settings and saved setlists as real files in this folder:

    config.json     colours, text size, Google Drive link (shared with the desktop app)
    setlists/*.json saved setlists (the same files the desktop app uses)
    JJ-SETLIST-TEMPLATE.xlsx  template spreadsheet and CSV made with Help > Create Template...
    JJ-SETLIST-TEMPLATE.csv

Start it with the launcher for your system - "Start JJ's Setlist.bat" (Windows),
"Start JJ's Setlist.command" (Mac) or "start-jjs-setlist.sh" (Linux) - or run
"python3 serve.py". Close its window to stop it.
Only this computer can connect (it listens on localhost). Needs Python 3.8 or
newer; standard library only.
"""
import sys

if sys.version_info < (3, 8):
    sys.exit("JJ's Setlist needs Python 3.8 or newer (this is Python %d.%d)." % sys.version_info[:2])

import http.server
import json
import os
import re
import socket
import socketserver
import subprocess
import threading
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from pathlib import Path

APP = "jjs-setlist"
VERSION = "1.5.B"
PORT = 8765                     # the address to bookmark: http://localhost:8765/
HOST = "127.0.0.1"
ROOT = Path(__file__).resolve().parent
WEB_DIR = ROOT / "web"
SETLIST_DIR = ROOT / "setlists"
CONFIG_FILE = ROOT / "config.json"
MAX_BODY = 5 * 1024 * 1024      # a setlist is a few KB; refuse anything silly


def safe_name(name):
    """A setlist file name (no extension) that can't escape the setlists folder."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", name).strip(" .")
    return name if name and name not in (".", "..") else ""


def spreadsheet_path(filename):
    """A spreadsheet directly in this folder (.xlsx / .xlsm / .csv), or None."""
    name = safe_name(filename)
    if not name or Path(name).suffix.lower() not in (".xlsx", ".xlsm", ".csv"):
        return None
    return ROOT / name


def unused_path(path):
    """path, or "name (2).ext", "name (3).ext"... if it's taken."""
    n = 2
    candidate = path
    while candidate.exists():
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        n += 1
    return candidate


def open_with_default_app(path):
    """Open a file with its usual program (e.g. Excel for .xlsx)."""
    if sys.platform == "win32":
        os.startfile(str(path))          # noqa: S606 - a spreadsheet in our own folder
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def read_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return default


def write_json(path, data):
    """Write via a temporary file, so a crash never leaves half a file."""
    tmp = path.with_name(path.name + ".part")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)


class Handler(http.server.SimpleHTTPRequestHandler):
    server_version = "JJsSetlist/" + VERSION

    # ---------- files: the app from web/, anything else from this folder
    # (so songsheet links like "3-temp-songs/My Song.pdf" open, as in the desktop app)
    def translate_path(self, path):
        rel = urllib.parse.unquote(urllib.parse.urlsplit(path).path).lstrip("/")
        for base in (WEB_DIR, ROOT):
            target = (base / rel).resolve()
            if (target == base or base in target.parents) and target.exists():
                return str(target)
        return str(WEB_DIR / "__missing__")

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")      # always the latest app files
        super().end_headers()

    def log_message(self, fmt, *args):                       # keep the window quiet
        if not self.path.startswith("/api/") and " 404 " not in (fmt % args):
            return
        sys.stderr.write("  " + (fmt % args) + "\n")

    # ---------- API
    def _allowed(self):
        """Only the app itself: requests to localhost, from our own page.
        Stops other websites (and DNS-rebinding tricks) using the API."""
        host = (self.headers.get("Host") or "").rsplit(":", 1)[0].strip("[]")
        if host not in ("localhost", "127.0.0.1", "::1"):
            return False
        origin = self.headers.get("Origin")
        if origin and urllib.parse.urlsplit(origin).hostname not in ("localhost", "127.0.0.1", "::1"):
            return False
        return self.headers.get("X-JJS") == "1"

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _raw_body(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_BODY:
            raise ValueError("too large")
        return self.rfile.read(length)

    def _body(self):
        return json.loads(self._raw_body().decode("utf-8") or "null")

    def _api(self, method):
        if not self._allowed():
            return self._send_json({"error": "not allowed"}, 403)
        parts = [urllib.parse.unquote(p) for p in urllib.parse.urlsplit(self.path).path.split("/")[2:]]
        try:
            if parts == ["ping"] and method == "GET":
                return self._send_json({"app": APP, "version": VERSION, "folder": str(ROOT)})

            if parts == ["config"]:
                if method == "GET":
                    return self._send_json(read_json(CONFIG_FILE, {}))
                if method == "PUT":
                    # Merge, so keys the desktop app keeps here (window size,
                    # printer...) are left alone.
                    updates = self._body()
                    if not isinstance(updates, dict):
                        raise ValueError("expected an object")
                    config = read_json(CONFIG_FILE, {})
                    config.update(updates)
                    write_json(CONFIG_FILE, config)
                    return self._send_json({"ok": True})

            if parts == ["setlists"] and method == "GET":
                found = {}
                for p in sorted(SETLIST_DIR.glob("*.json")):
                    data = read_json(p, None)
                    if isinstance(data, dict) and isinstance(data.get("sets"), list):
                        found[p.stem] = data
                return self._send_json(found)

            if len(parts) == 2 and parts[0] == "setlists":
                name = safe_name(parts[1])
                if not name:
                    return self._send_json({"error": "bad name"}, 400)
                path = SETLIST_DIR / f"{name}.json"
                if method == "PUT":
                    data = self._body()
                    if not isinstance(data, dict) or not isinstance(data.get("sets"), list):
                        raise ValueError("not a setlist")
                    SETLIST_DIR.mkdir(exist_ok=True)
                    write_json(path, data)
                    return self._send_json({"ok": True, "file": str(path)})
                if method == "DELETE":
                    path.unlink(missing_ok=True)
                    return self._send_json({"ok": True})

            # Help > Create Template: save a new spreadsheet in this folder,
            # never replacing one that's already here ("JJ-SETLIST-TEMPLATE (2).xlsx"...).
            if len(parts) == 2 and parts[0] == "template" and method == "PUT":
                path = spreadsheet_path(parts[1])
                if not path:
                    return self._send_json({"error": "bad name"}, 400)
                data = self._raw_body()
                if path.suffix.lower() != ".csv" and not data.startswith(b"PK"):
                    raise ValueError("not an Excel workbook")
                path = unused_path(path)
                path.write_bytes(data)
                return self._send_json({"ok": True, "file": path.name, "path": str(path)})

            # "Open it in Excel": open a spreadsheet in this folder with its usual program.
            if len(parts) == 2 and parts[0] == "open" and method == "PUT":
                path = spreadsheet_path(parts[1])
                if not path or not path.is_file():
                    return self._send_json({"error": "not found"}, 404)
                open_with_default_app(path)
                return self._send_json({"ok": True})
        except (ValueError, json.JSONDecodeError) as exc:
            return self._send_json({"error": f"bad request: {exc}"}, 400)
        except OSError as exc:
            return self._send_json({"error": str(exc)}, 500)
        return self._send_json({"error": "not found"}, 404)

    def do_GET(self):
        if self.path.startswith("/api/"):
            return self._api("GET")
        return super().do_GET()

    def do_PUT(self):
        return self._api("PUT") if self.path.startswith("/api/") else self.send_error(405)

    def do_DELETE(self):
        return self._api("DELETE") if self.path.startswith("/api/") else self.send_error(405)


class Server(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    # Windows: off, or two programs could share the port. Mac/Linux: on, so the
    # helper can restart at once after being closed (it can't share a port there).
    allow_reuse_address = sys.platform != "win32"

    def server_bind(self):
        # HTTPServer looks up the computer's network name here, which can take
        # several seconds on Windows - not needed for a localhost-only server.
        socketserver.TCPServer.server_bind(self)
        self.server_name, self.server_port = "localhost", self.server_address[1]


class Server6(Server):
    address_family = socket.AF_INET6     # "localhost" can mean ::1 as well as 127.0.0.1


def port_in_use(port):
    """True if any program answers on this port, on either loopback address."""
    # Short timeout: on Windows a refused localhost connection takes ~2 s to be
    # reported, while a program that is listening answers at once.
    for host in ("127.0.0.1", "::1"):
        try:
            with socket.create_connection((host, port), timeout=0.3):
                return True
        except OSError:
            pass
    return False


def already_running(port):
    """True if JJ's Setlist's own helper is already on this port."""
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{port}/api/ping", headers={"X-JJS": "1"})
        with urllib.request.urlopen(req, timeout=2) as r:
            return json.load(r).get("app") == APP
    except (OSError, ValueError):
        return False


def main():
    try:
        sys.stdout.reconfigure(line_buffering=True)   # messages appear straight away
    except AttributeError:
        pass
    if not (WEB_DIR / "index.html").is_file():
        print(f"Can't find web\\index.html next to serve.py (in {ROOT}).")
        return 1
    port = PORT
    if "--port" in sys.argv:                 # e.g. for testing: python serve.py --port 8770
        try:
            port = int(sys.argv[sys.argv.index("--port") + 1])
        except (IndexError, ValueError):
            print("Usage: python serve.py [--port NUMBER] [--no-browser]")
            return 1
    busy = (f"Port {port} is being used by another program (perhaps another web server).\n"
            "Close it and start JJ's Setlist again.")
    if port_in_use(port):
        if not already_running(port):
            print(busy)
            return 1
        print(f"JJ's Setlist is already running - opening http://localhost:{port}/")
        if "--no-browser" not in sys.argv:
            webbrowser.open(f"http://localhost:{port}/")
        return 0
    try:
        server = Server((HOST, port), Handler)
    except OSError:
        print(busy)
        return 1
    try:                                 # also on IPv6 loopback, when the computer has it
        server6 = Server6(("::1", port), Handler)
        threading.Thread(target=server6.serve_forever, daemon=True).start()
    except OSError:
        server6 = None
    url = f"http://localhost:{port}/"
    print("=" * 60)
    print(f"  JJ's Setlist is running:  {url}")
    print(f"  Settings and setlists are saved in:  {ROOT}")
    print("  Keep this window open while you use it.")
    print("  Close this window (or press Ctrl+C) to stop it.")
    print("=" * 60)
    if "--no-browser" not in sys.argv:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        if server6:
            server6.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
