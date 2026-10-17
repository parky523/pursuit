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


MAX_UPLOAD = 200 * 1024 * 1024  # per request


def parse_multipart(ctype, body):
    """multipart/form-data -> ({field: str}, [(filename, bytes)]). Uses email, since cgi is gone in 3.13."""
    from email.parser import BytesParser
    from email.policy import HTTP
    msg = BytesParser(policy=HTTP).parsebytes(b"Content-Type: " + ctype.encode() + b"\r\n\r\n" + body)
    fields, files = {}, []
    for part in msg.iter_parts():
        name = part.get_param("name", header="content-disposition")
        data = part.get_payload(decode=True) or b""
        if part.get_filename():
            if data:
                files.append((part.get_filename(), data))
        elif name:
            fields[name] = data.decode("utf-8", "replace")
    return fields, files


def safe_name(name, limit=60):
    import re
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._")
    return stem[:limit] or "file"


def save_uploads(folder, files, label, note, log_dir, dest_label):
    """Write files as YYYY-MM-DD_HHMMSS_<label|original>.<ext> and log each one to <inbox>/_uploads.jsonl."""
    import datetime as dt
    from pathlib import Path
    folder.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    saved = []
    for i, (orig, data) in enumerate(files, 1):
        p = Path(orig)
        ext = p.suffix.lower() or (mimetypes.guess_extension("application/octet-stream") or "")
        base = safe_name(label) if label else safe_name(p.stem)
        if label and len(files) > 1:
            base += f"_{i}"
        target = folder / f"{stamp}_{base}{ext}"
        n = 2
        while target.exists():
            target = folder / f"{stamp}_{base}_{n}{ext}"
            n += 1
        target.write_bytes(data)
        saved.append({"name": target.name, "path": str(target), "size": len(data), "dest": dest_label,
                      "note": note, "time": dt.datetime.now().isoformat(timespec="seconds")})
    with open(log_dir / "_uploads.jsonl", "a", encoding="utf-8") as f:
        for s in saved:
            f.write(json.dumps(s, ensure_ascii=False) + "\n")
    return saved


def recent_uploads(log_dir, limit=30):
    try:
        lines = (log_dir / "_uploads.jsonl").read_text(encoding="utf-8").splitlines()
    except OSError:
        return []
    out = []
    for i, line in enumerate(lines):
        try:
            d = json.loads(line)
            d["id"] = i
            out.append(d)
        except ValueError:
            pass
    return out[-limit:][::-1]


class PhoneServer:
    def __init__(self, app, base, inbox):
        self.app, self.base, self.inbox = app, base, inbox
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
                                   "purpose": "any maskable"}],
                        # Android "Share -> Pursuit" from gallery, WhatsApp, Chrome, files...
                        "share_target": {"action": f"/share?k={server.token}", "method": "POST",
                                         "enctype": "multipart/form-data",
                                         "params": {"title": "title", "text": "text", "url": "url",
                                                    "files": [{"name": "files",
                                                               "accept": ["image/*", "application/pdf", "*/*"]}]}},
                    }), "application/manifest+json")
                if u.path == "/api/inbox":
                    return self.send(200, json.dumps(recent_uploads(server.inbox)))
                if u.path == "/inbox-file":
                    try:  # only files that were logged as uploads can be fetched
                        item = next(x for x in recent_uploads(server.inbox, 10**6) if x["id"] == int(q["id"][0]))
                        from pathlib import Path
                        f = Path(item["path"])
                        ctype = mimetypes.guess_type(f.name)[0] or "application/octet-stream"
                        return self.send(200, f.read_bytes(), ctype, cache=True)
                    except Exception:
                        return self.send(404, "{}")
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
                if u.path in ("/api/upload", "/share"):
                    return self.upload(u.path == "/share")
                if u.path != "/api/do":
                    return self.send(404, "{}")
                try:
                    n = int(self.headers.get("Content-Length") or 0)
                    payload = json.loads(self.rfile.read(n) or b"{}")
                    out = server.on_main(lambda: server.app.phone_do(payload))
                    self.send(200, json.dumps(out, default=str))
                except Exception as ex:
                    self.send(500, json.dumps({"error": str(ex)}))

            def upload(self, shared):
                """Camera/file upload from the Inbox tab, or Android share-to-Pursuit."""
                try:
                    n = int(self.headers.get("Content-Length") or 0)
                    if n > MAX_UPLOAD:
                        return self.send(413, json.dumps({"error": "too big (max 200 MB per send)"}))
                    fields, files = parse_multipart(self.headers.get("Content-Type", ""), self.rfile.read(n))
                    note = fields.get("note", "").strip()
                    if shared:  # shared links/text arrive as fields, not files
                        text = "\n".join(x for x in (fields.get("title"), fields.get("text"), fields.get("url")) if x)
                        note = note or text[:300]
                        if text and not files:
                            files = [("shared_link.md", text.encode("utf-8"))]
                    if not files:
                        return self.send(400, json.dumps({"error": "no file received"}))
                    folder, label = server.on_main(lambda: server.app.upload_target(fields.get("dest")))
                    saved = save_uploads(folder, files, fields.get("label", "").strip(), note, server.inbox, label)
                    server.on_main(lambda: server.app.notify_upload(len(saved), label))
                    if shared:  # back into the app, on the Inbox tab
                        self.send_response(303)
                        self.send_header("Location", f"/?k={server.token}&shared={len(saved)}#inbox")
                        self.end_headers()
                        return
                    self.send(200, json.dumps({"saved": saved}))
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
