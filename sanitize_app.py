#local interface for sanitize_pptx.py: python sanitize_app.py opens it in the browser.
#serves only this computer (127.0.0.1): decks and keys go from the browser to this script and back, nowhere else.
#slide previews are exported with PowerPoint when it is installed (Windows), otherwise the page works without them.
#the script keeps no deck: the page sends the open deck with every request.
#the desktop shortcut opens sanitize_desktop.py instead; this page is the larger view, with slide previews.

import base64
import io
import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import sanitize_pptx as sp

HERE = Path(__file__).resolve().parent
PAGE = HERE / "app" / "index.html"
SECRET = secrets.token_urlsafe(16) #only the page this script serves knows it, so other sites can't call the API
MAX_SLIDES = 60 #previews for the first slides only, exporting is slow on long decks

#PowerPoint exports each slide as a PNG; it is closed again only if this script started it
PS_EXPORT = r"""
$ErrorActionPreference = 'Stop'
$running = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
try {
  $p = $app.Presentations.Open($env:SAN_DECK, -1, 0, 0)
  try {
    $w = 480
    $h = [int]($w * $p.PageSetup.SlideHeight / $p.PageSetup.SlideWidth)
    $n = [Math]::Min($p.Slides.Count, [int]$env:SAN_MAX)
    for ($i = 1; $i -le $n; $i++) { $p.Slides.Item($i).Export((Join-Path $env:SAN_OUT ('{0:D3}.png' -f $i)), 'PNG', $w, $h) }
  } finally { $p.Close() }
} finally { if (-not $running) { $app.Quit() } }
"""


def thumbnails(deck):
    #slide images as data URLs, or None when PowerPoint can't export them
    if sys.platform != "win32":
        return None
    tmp = Path(tempfile.mkdtemp(prefix="sanitize_"))
    try:
        (tmp / "deck.pptx").write_bytes(deck)
        env = {**os.environ, "SAN_DECK": str(tmp / "deck.pptx"), "SAN_OUT": str(tmp), "SAN_MAX": str(MAX_SLIDES)}
        script = base64.b64encode(PS_EXPORT.encode("utf-16-le")).decode()
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", script], env=env,
                       capture_output=True, timeout=300, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        pngs = sorted(tmp.glob("*.png"))
        return ["data:image/png;base64," + base64.b64encode(p.read_bytes()).decode() for p in pngs] or None
    except (OSError, subprocess.SubprocessError):
        return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True) #the temp copy holds the client's deck


def slide_titles(deck):
    #{slide number: title text}, for the list of slides to check
    z = zipfile.ZipFile(io.BytesIO(deck))
    titles = {}
    for part, loc in sp.build_locations(z).items():
        n = sp.slide_no(loc)
        if n and sp.SLIDE_RX.match(part):
            root = sp.parse_root(z.read(part))
            keep = sp.title_paragraphs(root)
            paras = [p for p in root.getElementsByTagNameNS(sp.NS_A, "p") if p in keep]
            text = " ".join("".join(sp.get_text(t) for t in p.getElementsByTagNameNS(sp.NS_A, "t")) for p in paras)
            titles[n] = re.sub("[⁠​‌⁣]", "", text).strip()
    return titles


def slides_in(locations):
    #slide numbers named in report locations ("Slide 3", "Slide 3 chart"...), in order
    return sorted({int(m.group(1)) for loc in locations for m in [re.match(r"Slide (\d+)", loc or "")] if m})


def b64(data):
    return base64.b64encode(data).decode()


def stem(name):
    return name[:-5] if name.lower().endswith(".pptx") else name


def api_open(req):
    data = base64.b64decode(req["deck"])
    sp.check_deck(data)
    return {"name": req["name"], "slides": thumbnails(data)}


def api_preview(req):
    full = req["mode"] == "all"
    r = sp.preview_deck(base64.b64decode(req["deck"]), None if full else req.get("keywords", ""), full=full)
    found = [f for f in r["findings"] if not f["type"].startswith("METADATA_")] #file properties are always cleared
    return {"count": len(found), "marked": slides_in(f["location"] for f in found),
            "slides": thumbnails(r["deck"]), "deck": b64(r["deck"]), "summary": r["summary"]["text"]}


def api_mask(req):
    data, name = base64.b64decode(req["deck"]), req["name"]
    if req["mode"] == "all":
        r = sp.mask_all(data, name, restore_hint="Open the returned deck, choose Restore and pick this key.")
    else:
        r = sp.mask_keywords(data, req.get("keywords", ""), name,
                             restore_hint="Open the returned deck, choose Restore and pick this key.")
    deck_name, key_name = stem(name) + "_redacted.pptx", stem(name) + "_restore_key.xlsx"
    return {"deck": b64(r["deck"]), "deck_name": deck_name, "key": b64(r["key"]), "key_name": key_name,
            "slides": thumbnails(r["deck"]), "notes": r["notes"], "messages": r["messages"], "summary": r["summary"]["text"]}


def api_restore(req):
    r = sp.restore_deck(base64.b64decode(req["deck"]), base64.b64decode(req["key"])) #never forced: another deck's key writes wrong text
    name = stem(req["name"]) + "_restored.pptx"
    layout = next((m["items"] for m in r["messages"] if m["code"] == "check_layout"), [])
    titles = slide_titles(r["deck"])
    return {"deck": b64(r["deck"]), "name": name, "slides": thumbnails(r["deck"]),
            "check": [{"slide": n, "title": titles.get(n, "")} for n in slides_in(layout)],
            "messages": [m for m in r["messages"] if m["code"] != "check_layout"], "summary": r["summary"]["text"],
            "restored": r["summary"]["restored"], "to_restore": r["summary"]["to_restore"]}


ROUTES = {"/api/open": api_open, "/api/preview": api_preview, "/api/mask": api_mask, "/api/restore": api_restore}
FRIENDLY = {"wrong_key": "This key was made for a different deck. Pick the key saved when this deck was masked."}


class Handler(BaseHTTPRequestHandler):
    def send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else json.dumps(body).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self.send(200, PAGE.read_text(encoding="utf-8").replace("{{SECRET}}", SECRET).encode("utf-8"),
                      "text/html; charset=utf-8")
        else:
            self.send(404, {"error": "not_found", "message": "Not found"})

    def do_POST(self):
        route = ROUTES.get(self.path)
        if route is None or self.headers.get("X-Sanitize") != SECRET:
            return self.send(404, {"error": "not_found", "message": "Not found"})
        try:
            req = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))))
            if not req.get("deck"):
                return self.send(400, {"error": "no_deck", "message": "Open a deck first."})
            self.send(200, route(req))
        except sp.SanitizeError as e:
            self.send(400, {"error": e.code, "message": FRIENDLY.get(e.code, str(e))})
        except Exception as e: #shown in the page instead of a silent failure
            self.send(500, {"error": "failed", "message": f"Something went wrong: {e}"})

    def log_message(self, *args): #keep file names and paths out of the console
        pass


def main():
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler) #this computer only, on a free port
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"Sanitize is running at {url} (this computer only). Close this window or press Ctrl+C to stop.")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
