"""Pursuit phone server: serves phone.html + a small JSON API from inside the desktop app.

Every request is checked against a secret key (?k=... or X-Key header). The key lives in checklist.json.
All app logic runs on the Tk main thread via app.calls, so the phone and the desktop window share one
state and every phone edit lands in the same Excel COM path as the desktop buttons.
"""
import json
import mimetypes
import secrets
import socket
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

PORT = 8787
STATIC = {"/icon-192.png", "/icon-512.png", "/apple-touch-icon.png", "/favicon.png"}


def local_ips():
    """(lan, tailscale) IPv4 addresses of this PC; tailscale = 100.64.0.0/10."""
    ips = set()
    try:
        ips.update(socket.gethostbyname_ex(socket.gethostname())[2])
    except OSError:
        pass
    try:  # the default-route address, in case the hostname lookup misses it
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ips.add(s.getsockname()[0])
        s.close()
    except OSError:
        pass
    ts = [ip for ip in ips if ip.startswith("100.") and 64 <= int(ip.split(".")[1]) <= 127]
    lan = [ip for ip in ips if ip not in ts and not ip.startswith("127.")]
    return (lan[0] if lan else None), (ts[0] if ts else None)


class PhoneServer:
    def __init__(self, app, base):
        self.app, self.base = app, base
        if not app.s.get("token"):
            app.s["token"] = secrets.token_urlsafe(9)
        self.token = app.s["token"]
        self.error = ""

    def urls(self):
        lan, ts = local_ips()
        q = f":{PORT}/?k={self.token}"
        return {"lan": f"http://{lan}{q}" if lan else None, "tailscale": f"http://{ts}{q}" if ts else None,
                "https": tailscale_https(self.token)}


    def on_main(self, fn):
        """Run fn on the Tk thread and wait for its result."""
        ev, box = threading.Event(), {}
        self.app.calls.put((fn, ev, box))
        if not ev.wait(15):
            raise TimeoutError("desktop app busy")
        if "error" in box:
            raise box["error"]
        return box.get("result")

    def start(self):
        server = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def send(self, code, body, ctype="application/json", cache=False):
                data = body if isinstance(body, bytes) else body.encode("utf-8")
                self.send_response(code)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "public, max-age=86400" if cache else "no-store")
                self.end_headers()
                self.wfile.write(data)

            def authed(self, q):
                key = self.headers.get("X-Key") or (q.get("k") or [""])[0]
                return secrets.compare_digest(key, server.token)

            def do_GET(self):
                u = urlparse(self.path)
                q = parse_qs(u.query)
                if u.path in STATIC:
                    name = "apple-touch-icon.png" if u.path == "/favicon.png" else u.path[1:]
                    return self.send(200, (server.base / name).read_bytes(), "image/png", cache=True)
                if u.path == "/sw.js":
                    return self.send(200, (server.base / "sw.js").read_bytes(), "text/javascript")
                if u.path.startswith("/fonts/") and u.path.endswith(".ttf") and "/.." not in u.path:
                    f = server.base / "fonts" / u.path.rsplit("/", 1)[1]
                    if f.exists():
                        return self.send(200, f.read_bytes(), "font/ttf", cache=True)
                if not self.authed(q):
                    return self.send(403, "<h1 style='font-family:monospace'>Pursuit: wrong or missing key</h1>",
                                     "text/html")
                if u.path == "/":
                    html = (server.base / "phone.html").read_text(encoding="utf-8")
                    return self.send(200, html.replace("{{TOKEN}}", server.token), "text/html; charset=utf-8")
                if u.path == "/manifest.webmanifest":
                    return self.send(200, json.dumps({
                        "name": "Pursuit", "short_name": "Pursuit", "start_url": f"/?k={server.token}",
                        "scope": "/", "display": "standalone", "background_color": "#0b0b0c",
                        "theme_color": "#0b0b0c", "orientation": "portrait",
                        "icons": [{"src": "/icon-192.png", "sizes": "192x192", "type": "image/png",
                                   "purpose": "any maskable"},
                                  {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png",
                                   "purpose": "any maskable"}]}), "application/manifest+json")
                if u.path == "/api/state":
                    try:
                        return self.send(200, json.dumps(server.on_main(server.app.state_json), default=str))
                    except Exception as ex:
                        return self.send(500, json.dumps({"error": str(ex)}))
                self.send(404, "{}")

            def do_POST(self):
                u = urlparse(self.path)
                if not self.authed(parse_qs(u.query)):
                    return self.send(403, "{}")
                if u.path != "/api/do":
                    return self.send(404, "{}")
                try:
                    n = int(self.headers.get("Content-Length") or 0)
                    payload = json.loads(self.rfile.read(n) or b"{}")
                    out = server.on_main(lambda: server.app.phone_do(payload))
                    self.send(200, json.dumps(out, default=str))
                except Exception as ex:
                    self.send(500, json.dumps({"error": str(ex)}))

        try:
            httpd = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
        except OSError as ex:
            self.error = f"phone server off: port {PORT} busy ({ex.strerror})"
            return False
        threading.Thread(target=httpd.serve_forever, daemon=True).start()
        return True


TAILSCALE = r"C:\Program Files\Tailscale\tailscale.exe"


def tailscale_https(token):
    """https://<pc>.<tailnet>.ts.net/?k=... when `tailscale serve --bg 8787` is active - needed for app install."""
    import subprocess
    try:
        run = lambda *a: subprocess.run([TAILSCALE, *a], capture_output=True, text=True, timeout=5,
                                        creationflags=0x08000000).stdout  # CREATE_NO_WINDOW
        if str(PORT) not in run("serve", "status"):
            return None
        name = json.loads(run("status", "--json"))["Self"]["DNSName"].rstrip(".")
        return f"https://{name}/?k={token}"
    except Exception:
        return None
