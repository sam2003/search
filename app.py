#!/usr/bin/env python3
"""
Registry Search — LAN document search over many same-format Excel exports.

Run on one PC:      python app.py
Open on any PC:     http://<host-pc-ip>:5000

Built for an old server. Between requests the server holds file paths and
nothing else. During a search it opens one file at a time, streams its rows,
keeps only the rows that match, and releases the file before moving on.

Everything you'd normally want to change lives in the CONFIG block below.
"""

import io
import os
import re
import csv
import uuid
import codecs
import socket
import shutil
import tempfile
import threading
from datetime import datetime

from lxml import etree
from flask import (Flask, Request, request, jsonify, render_template_string,
                   Response, make_response, stream_with_context)

# ============================================================================
# CONFIG — edit this block only
# ============================================================================

# Hosting platforms hand the port over in an environment variable; locally
# these defaults apply and nothing needs setting.
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "5000"))

# Largest single upload accepted. Uploads stream to disk and never sit in RAM,
# so this guards disk rather than memory. Only matters on a public URL.
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "300"))

# Safety limit on rows held in RAM during one search. This is the only thing
# in the program whose memory cost grows with your data. Measured on this data:
# about 5 MB of peak server RAM per 1000 matched rows, so 2000 is roughly
# +10 MB. Raise it if the server has headroom.
MAX_MATCHES = 2000

# Where uploaded files are parked on the server's disk. Cleared on startup.
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "registry-search-files")

READ_CHUNK = 64 * 1024

# Rows that appear identically in more than one file are kept by default.
# Set to True to collapse them instead.
MERGE_DUPLICATES = False

# Columns offered as filters, in the order they appear in the UI. Each is a
# case-insensitive "contains" box.
#   loose: True -> ignores spaces, quotes and braces when matching (good for
#                  Marathi party names that are exported without spaces)
FILTERS = [
    {"col": "docno",               "label": "Document no",       "type": "text"},
    {"col": "docname",             "label": "Document type",     "type": "text"},
    {"col": "sellerparty",         "label": "Seller",            "type": "text", "loose": True},
    {"col": "purchaserparty",      "label": "Purchaser",         "type": "text", "loose": True},
    {"col": "propertydescription", "label": "Property contains", "type": "text", "loose": True},
]

# Columns shown in the results table, in order. Leave empty to show everything.
DISPLAY_COLUMNS = [
    "docno", "registrationdate", "docname", "areaname", "sroname",
    "sellerparty", "purchaserparty", "propertydescription",
]

# Columns rendered as long text with a "more" toggle instead of one line.
LONG_TEXT_COLUMNS = ["propertydescription", "sellerparty", "purchaserparty"]

FILE_EXTENSIONS = (".xls", ".xlsx", ".xlsm", ".csv", ".htm", ".html")

# AdSense publisher id, Formspree contact endpoint, and the address the form
# emails. Override any of these via environment variables on the host.
ADSENSE_CLIENT = os.environ.get("ADSENSE_CLIENT", "ca-pub-5876158560001890")
CONTACT_EMAIL = os.environ.get("CONTACT_EMAIL", "piratehere123@gmail.com")

# ============================================================================
# Small helpers
# ============================================================================

_LOOSE_RE = re.compile(r'[\s"\'{}\[\]().,\-_/]')


def loose(s):
    return _LOOSE_RE.sub("", s).lower()


def clean(s):
    return " ".join(s.split()) if s else ""


# ============================================================================
# Streaming readers — each yields the header list, then row lists
# ============================================================================

def _text_chunks(path):
    """Yield utf-8 bytes from a file of unknown encoding, without ever holding
    more than one chunk. Handles the UTF-16 exports from the registry portal."""
    with open(path, "rb") as fh:
        head = fh.read(4)
        fh.seek(0)
        if head[:2] in (b"\xff\xfe", b"\xfe\xff"):
            enc = "utf-16"
        else:
            enc = "utf-8"
        dec = codecs.getincrementaldecoder(enc)("replace")
        while True:
            chunk = fh.read(READ_CHUNK)
            if not chunk:
                break
            text = dec.decode(chunk)
            if text:
                yield text.encode("utf-8")
        tail = dec.decode(b"", True)
        if tail:
            yield tail.encode("utf-8")


def iter_html_rows(path):
    """Stream <tr> elements out of an HTML table, dropping each one as soon as
    it has been read so the tree never grows."""
    parser = etree.HTMLPullParser(events=("end",), tag="tr", recover=True,
                                  encoding="utf-8")

    def drain():
        for _event, el in parser.read_events():
            cells = [c for c in el if c.tag in ("td", "th")]
            if cells:
                yield [clean("".join(c.itertext())) for c in cells]
            parent = el.getparent()
            el.clear()
            if parent is not None:
                while el.getprevious() is not None:
                    del parent[0]

    for chunk in _text_chunks(path):
        parser.feed(chunk)
        for row in drain():
            yield row
    try:
        parser.close()
    except etree.XMLSyntaxError:
        pass
    for row in drain():
        yield row


def iter_csv_rows(path):
    with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
        for row in csv.reader(fh):
            yield [clean(c) for c in row]


def iter_xlsx_rows(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(values_only=True):
            yield [clean(str(c)) if c is not None else "" for c in row]
    finally:
        wb.close()


def iter_xls_rows(path):
    """Real OLE2 .xls. xlrd loads the whole book, so this is the one format
    that isn't streamed — rare for these exports."""
    import xlrd
    book = xlrd.open_workbook(path)
    try:
        sheet = book.sheet_by_index(0)
        for i in range(sheet.nrows):
            yield [clean(str(v)) for v in sheet.row_values(i)]
    finally:
        book.release_resources()


def iter_rows(path):
    """Pick a reader by sniffing the first bytes, not by extension — the
    portal's .xls files are really UTF-16 HTML tables."""
    with open(path, "rb") as fh:
        head = fh.read(8)

    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return iter_xls_rows(path)
    if head[:2] == b"PK":
        return iter_xlsx_rows(path)
    if path.lower().endswith(".csv") and head[:2] not in (b"\xff\xfe", b"\xfe\xff"):
        return iter_csv_rows(path)
    return iter_html_rows(path)


# ============================================================================
# Filters compiled to per-file predicates
# ============================================================================

def cell(row, i):
    return row[i] if i < len(row) else ""


def build_predicates(params, header):
    """Turn the filter form into a list of tests bound to this file's column
    positions. Blank inputs produce no test, so every combination works."""
    idx = {c: i for i, c in enumerate(header)}
    preds = []

    free = str(params.get("_all", "")).strip()
    if free:
        needle = loose(free)
        preds.append(lambda row, n=needle: n in loose(" ".join(row)))

    for f in FILTERS:
        col = f["col"]
        i = idx.get(col)
        if i is None:
            continue

        if f["type"] == "text":
            val = str(params.get(col, "")).strip()
            if not val:
                continue
            if f.get("loose"):
                preds.append(lambda row, i=i, n=loose(val): n in loose(cell(row, i)))
            else:
                preds.append(lambda row, i=i, n=val.lower(): n in cell(row, i).lower())

    return preds


# ============================================================================
# The search itself
# ============================================================================

def run_search(paths, params, dedupe=False, cap=None):
    """One file at a time: open, stream, keep matches, close, free.
    Only matched rows survive between files."""
    cap = cap or MAX_MATCHES
    matches, notes = [], []
    seen = set()
    scanned = 0
    truncated = False
    columns = None

    for path in paths:
        name = os.path.basename(path)
        rows_here = 0
        try:
            it = iter_rows(path)
            header = None
            for raw in it:
                if header is None:
                    header = [c.strip().lower() for c in raw]
                    keep = [c for c in DISPLAY_COLUMNS if c in header] or header
                    if columns is None:
                        columns = keep
                    take = [header.index(c) if c in header else -1 for c in columns]
                    preds = build_predicates(params, header)
                    continue

                rows_here += 1
                scanned += 1
                if all(p(raw) for p in preds):
                    if dedupe:
                        h = hash(tuple(raw))
                        if h in seen:
                            continue
                        seen.add(h)
                    if len(matches) >= cap:
                        truncated = True
                        break
                    matches.append([cell(raw, i) if i >= 0 else "" for i in take])

            notes.append({"file": name, "rows": rows_here, "ok": True})
        except Exception as exc:
            notes.append({"file": name, "rows": 0, "ok": False,
                          "error": f"{type(exc).__name__}: {exc}"})
        finally:
            it = None

        if truncated:
            break

    return {
        "columns": columns or [],
        "matches": matches,
        "scanned": scanned,
        "truncated": truncated,
        "notes": notes,
    }


# ============================================================================
# App — sessions hold paths only
# ============================================================================

class DiskRequest(Request):
    """Send every uploaded file straight to disk. Nothing is buffered in RAM."""

    def _get_file_stream(self, total_content_length, content_type,
                         filename=None, content_length=None):
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        return tempfile.NamedTemporaryFile(dir=UPLOAD_DIR, prefix="part-",
                                           suffix=".tmp", delete=False)


app = Flask(__name__)
app.request_class = DiskRequest
# Marathi text as raw UTF-8 rather than \uXXXX escapes — roughly halves the
# size of every response, in RAM and on the wire.
app.json.ensure_ascii = False
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

# Behind a platform's HTTPS proxy, trust the forwarded scheme.
if os.environ.get("BEHIND_PROXY", "1") == "1":
    from werkzeug.middleware.proxy_fix import ProxyFix
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

_lock = threading.Lock()
SESSIONS = {}   # sid -> {"paths": [...], "dir": str}
MAX_SESSIONS = 40


def get_sid(req):
    return req.cookies.get("xls_sid") or uuid.uuid4().hex


def set_sid_cookie(resp, sid):
    resp.set_cookie("xls_sid", sid, max_age=60 * 60 * 24 * 30, samesite="Lax",
                    httponly=True, secure=request.is_secure)
    return resp


def drop_session(sid):
    SESSIONS.pop(sid, None)
    shutil.rmtree(os.path.join(UPLOAD_DIR, sid), ignore_errors=True)


def set_session_paths(sid, new_paths, updir):
    """A fresh selection REPLACES the previous one. Old uploaded copies are
    deleted so a removed set can't linger in later searches."""
    with _lock:
        old = SESSIONS.get(sid)
        if old and old.get("dir") and old["dir"] != updir:
            shutil.rmtree(old["dir"], ignore_errors=True)
        SESSIONS[sid] = {"paths": list(new_paths), "dir": updir,
                         "touched": datetime.now()}
        while len(SESSIONS) > MAX_SESSIONS:
            oldest = min(SESSIONS, key=lambda k: SESSIONS[k]["touched"])
            drop_session(oldest)


def get_session(sid):
    with _lock:
        s = SESSIONS.get(sid)
        if s:
            s["touched"] = datetime.now()
        return s


def save_upload(fs, dest):
    """Move the spooled temp file into place — no copy through RAM."""
    stream = getattr(fs, "stream", None)
    src = getattr(stream, "name", None)
    if src and os.path.exists(str(src)):
        try:
            stream.close()
        except OSError:
            pass
        os.replace(src, dest)
        return
    with open(dest, "wb") as out:
        shutil.copyfileobj(fs.stream, out, READ_CHUNK)


@app.route("/ads.txt")
def ads_txt():
    return Response(
        "google.com, pub-5876158560001890, DIRECT, f08c47fec0942fa0\n",
        mimetype="text/plain",
    )


@app.route("/healthz")
def healthz():
    """Cheap endpoint for an uptime pinger. Keeping this warm every 10 minutes
    stops a free-tier host from sleeping and wiping everyone's uploads."""
    return "ok", 200


@app.route("/")
def index():
    resp = make_response(render_template_string(PAGE, filters=FILTERS, cap=MAX_MATCHES))
    if not request.cookies.get("xls_sid"):
        resp.set_cookie("xls_sid", uuid.uuid4().hex, max_age=60 * 60 * 24 * 30, samesite="Lax")
    return resp


@app.route("/api/select", methods=["POST"])
def api_select():
    """Record where the files are. Nothing is read or parsed here."""
    sid = get_sid(request)
    # A new upload replaces the old set, so give it its own directory; the
    # previous one is deleted in set_session_paths.
    updir = os.path.join(UPLOAD_DIR, sid, uuid.uuid4().hex)
    os.makedirs(updir, exist_ok=True)
    paths = []


    for fs in request.files.getlist("files"):
        base = os.path.basename(fs.filename or "")
        if not base or base.startswith("~$") or not base.lower().endswith(FILE_EXTENSIONS):
            continue
        dest = os.path.join(updir, f"{len(paths):04d}-{base}")
        save_upload(fs, dest)
        paths.append(dest)

    if not paths:
        return jsonify({"error": "Nothing usable in that selection — looking for "
                                 + ", ".join(FILE_EXTENSIONS) + "."}), 400

    set_session_paths(sid, paths, updir)
    all_paths = paths
    return set_sid_cookie(jsonify({
        "files": len(all_paths),
        "names": [os.path.basename(p).split("-", 1)[-1] for p in all_paths][:400],
    }), sid)


@app.route("/api/search", methods=["POST"])
def api_search():
    sid = get_sid(request)
    sess = get_session(sid)
    if not sess or not sess.get("paths"):
        return jsonify({"error": "No files selected — if they were there a moment "
                                 "ago, the server restarted and cleared them. "
                                 "Drop them in again."}), 400

    params = request.get_json(silent=True) or {}
    paths = sess["paths"]
    out = run_search(paths, params, dedupe=MERGE_DUPLICATES)

    return jsonify({
        "total": len(out["matches"]),
        "scanned": out["scanned"],
        "files": len(paths),
        "truncated": out["truncated"],
        "cap": MAX_MATCHES,
        "columns": out["columns"],
        "long": [c for c in LONG_TEXT_COLUMNS if c in out["columns"]],
        "notes": out["notes"],
        "rows": out["matches"],
    })


@app.route("/api/export", methods=["POST"])
def api_export():
    """Streamed straight to the browser — matched rows are written out as they
    are found rather than collected first."""
    sid = get_sid(request)
    sess = get_session(sid)
    if not sess:
        return jsonify({"error": "No files selected."}), 400

    params = request.get_json(silent=True) or {}
    paths = sess.get("paths") or []
    dedupe = MERGE_DUPLICATES

    def generate():
        buf = io.StringIO()
        w = csv.writer(buf)
        yield "\ufeff"
        seen = set()
        header_written = False

        for path in paths:
            name = os.path.basename(path)
            try:
                header = None
                for raw in iter_rows(path):
                    if header is None:
                        header = [c.strip().lower() for c in raw]
                        cols = [c for c in DISPLAY_COLUMNS if c in header] or header
                        take = [header.index(c) for c in cols]
                        preds = build_predicates(params, header)
                        if not header_written:
                            w.writerow(cols)
                            header_written = True
                        continue
                    if all(p(raw) for p in preds):
                        if dedupe:
                            h = hash(tuple(raw))
                            if h in seen:
                                continue
                            seen.add(h)
                        w.writerow([cell(raw, i) for i in take])
                        if buf.tell() > 32768:
                            yield buf.getvalue()
                            buf.seek(0)
                            buf.truncate(0)
            except Exception:
                continue
        yield buf.getvalue()

    stamp = datetime.now().strftime("%Y%m%d-%H%M")
    return Response(
        stream_with_context(generate()),
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="results-{stamp}.csv"'},
    )


@app.route("/api/clear", methods=["POST"])
def api_clear():
    sid = get_sid(request)
    with _lock:
        drop_session(sid)
    return jsonify({"ok": True})


# ============================================================================
# Page
# ============================================================================


# ============================================================================
# Shared content-page building blocks (nav, footer, styling, AdSense)
# ============================================================================

ADSENSE_HEAD = (
    '<script async src="https://pagead2.googlesyndication.com/pagead/js/'
    'adsbygoogle.js?client=%s" crossorigin="anonymous"></script>'
) % ADSENSE_CLIENT


def ad_unit(slot_label=""):
    """A responsive AdSense display unit. Only rendered on content pages,
    beside real text — never on the search tool."""
    return (
        '<div class="adwrap"><span class="adtag">Advertisement</span>'
        '<ins class="adsbygoogle" style="display:block" '
        'data-ad-client="%s" data-ad-format="auto" '
        'data-full-width-responsive="true"></ins>'
        '<script>(adsbygoogle=window.adsbygoogle||[]).push({});</script></div>'
    ) % ADSENSE_CLIENT


NAV = r"""
<header class="topnav">
  <a class="brand" href="/">Registry<span class="amp">Search</span></a>
  <nav>
    <a href="/"{home}>Search</a>
    <a href="/guide"{guide}>Guide</a>
    <a href="/about"{about}>About IGR</a>
    <a href="/contact"{contact}>Contact</a>
    <a href="/privacy"{privacy}>Privacy</a>
  </nav>
</header>
"""


FOOTER = r"""
<footer class="sitefoot">
  <div class="footcol">
    <div class="footbrand">RegistrySearch</div>
    <p>A search tool for Maharashtra property-registration (IGR) records,
       built to look across many registry exports at once.</p>
  </div>
  <div class="footcol">
    <div class="foothead">Pages</div>
    <a href="/">Registry Search</a>
    <a href="/guide">How-to Guide</a>
    <a href="/about">About IGR &amp; This Tool</a>
    <a href="/contact">Contact</a>
    <a href="/privacy">Privacy Policy</a>
  </div>
  <div class="footcol">
    <div class="foothead">Note</div>
    <p>IGR registration records are public documents published by the
       Department of Registration &amp; Stamps, Government of Maharashtra.
       This tool helps search exports of that public data.</p>
  </div>
  <div class="footbar">RegistrySearch · search Maharashtra IGR records</div>
</footer>
"""

# CSS shared by the content pages (guide / about / contact / privacy)
CONTENT_CSS = r"""
  :root{
    --ink:#111827;--ink-2:#374151;--muted:#6b7280;--paper:#f9fafb;--sheet:#fff;
    --well:#f3f4f6;--line:#e5e7eb;--navy:#1e3a5f;--navy-2:#2563eb;--navy-3:#1d4ed8;
    --navy-tint:#e0eaf7;--navy-hover:#eff6ff;--navy-soft:#93b4d8;
    --sans:"Segoe UI",Roboto,Inter,system-ui,-apple-system,sans-serif;
    --serif:Georgia,"Times New Roman",serif;--mono:ui-monospace,Consolas,monospace;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--paper);color:var(--ink);font-family:var(--sans);
       font-size:16px;line-height:1.7;-webkit-font-smoothing:antialiased}

  .topnav{background:var(--navy);color:#fff;display:flex;align-items:center;
          gap:26px;padding:0 26px;flex-wrap:wrap;border-bottom:3px solid var(--navy-2)}
  .topnav .brand{font-family:var(--serif);font-size:20px;font-weight:700;color:#fff;
                 text-decoration:none;padding:16px 0;margin-right:auto}
  .topnav .brand .amp{color:var(--navy-soft);font-style:italic}
  .topnav nav{display:flex;gap:4px;flex-wrap:wrap}
  .topnav nav a{color:#cfe0f5;text-decoration:none;font-size:14px;font-weight:600;
                padding:16px 14px;border-bottom:3px solid transparent;margin-bottom:-3px}
  .topnav nav a:hover{color:#fff;background:rgba(255,255,255,.08)}
  .topnav nav a.active{color:#fff;border-bottom-color:#fff}

  .wrap{max-width:820px;margin:0 auto;padding:34px 24px 10px}
  .page-h1{font-family:var(--serif);font-size:32px;font-weight:700;line-height:1.2;
            margin:0 0 6px;color:var(--navy)}
  .lede{font-size:18px;color:var(--ink-2);margin:0 0 8px}
  .updated{font-family:var(--mono);font-size:12px;color:var(--muted);margin-bottom:26px}
  .wrap h2{font-family:var(--serif);font-size:24px;color:var(--navy);
           margin:34px 0 10px;padding-bottom:6px;border-bottom:2px solid var(--line)}
  .wrap h3{font-size:18px;color:var(--ink);margin:24px 0 6px}
  .wrap p{margin:0 0 14px}
  .wrap ul,.wrap ol{margin:0 0 16px;padding-left:22px}
  .wrap li{margin:0 0 7px}
  .wrap a{color:var(--navy-2);text-decoration:underline;text-underline-offset:2px}
  .mr{font-size:15px;color:var(--ink-2);background:var(--well);border-left:3px solid var(--navy-soft);
      padding:10px 14px;margin:0 0 16px;border-radius:0 4px 4px 0}
  .mr .lbl{font-family:var(--mono);font-size:10px;text-transform:uppercase;
           letter-spacing:.1em;color:var(--muted);display:block;margin-bottom:3px}
  table.fields{width:100%;border-collapse:collapse;margin:6px 0 20px;font-size:15px}
  table.fields th{background:var(--navy);color:#fff;text-align:left;padding:9px 12px;
                  font-size:12px;text-transform:uppercase;letter-spacing:.05em}
  table.fields td{padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top}
  table.fields tr:nth-child(even) td{background:var(--well)}
  table.fields td:first-child{font-family:var(--mono);font-size:13px;color:var(--navy);
                              white-space:nowrap;font-weight:600}
  .cta{display:inline-block;background:var(--navy-2);color:#fff;text-decoration:none;
       font-weight:700;padding:11px 22px;border-radius:7px;margin:6px 0 10px}
  .cta:hover{background:var(--navy-3)}
  .emaillink{font-size:18px;font-weight:700;color:var(--navy-2);
             text-decoration:none;border-bottom:2px solid var(--navy-2)}
  .emaillink:hover{color:var(--navy-3);border-bottom-color:var(--navy-3)}
  .callout{background:var(--navy-tint);border:1px solid #bfdbfe;border-radius:8px;
           padding:16px 18px;margin:18px 0}
  .callout .h{font-weight:700;color:var(--navy);margin-bottom:4px}

  .adwrap{margin:26px 0;text-align:center;min-height:100px;
          background:var(--well);border:1px solid var(--line);border-radius:8px;
          padding:10px;overflow:hidden}
  .adtag{display:block;font-family:var(--mono);font-size:10px;text-transform:uppercase;
         letter-spacing:.1em;color:var(--muted);margin-bottom:6px;text-align:left}

  form.contact{background:var(--sheet);border:1px solid var(--line);border-radius:10px;
               padding:24px;margin:10px 0 20px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
  form.contact label{display:block;font-size:12px;font-family:var(--mono);
                     text-transform:uppercase;letter-spacing:.06em;color:var(--muted);
                     font-weight:600;margin:14px 0 5px}
  form.contact input,form.contact textarea{width:100%;padding:11px 13px;
    border:1.5px solid var(--line);border-radius:6px;font-family:inherit;font-size:15px;
    color:var(--ink);background:#fff;outline:none}
  form.contact input:focus,form.contact textarea:focus{border-color:var(--navy-2);
    box-shadow:0 0 0 3px rgba(37,99,235,.1)}
  form.contact textarea{min-height:130px;resize:vertical}
  form.contact .req{color:#b3261e}
  form.contact button{margin-top:18px;background:var(--navy-2);color:#fff;border:0;
    font-size:15px;font-weight:700;padding:12px 26px;border-radius:7px;cursor:pointer}
  form.contact button:hover{background:var(--navy-3)}
  .formnote{font-size:13px;color:var(--muted);margin-top:10px}
  .flashmsg{padding:12px 16px;border-radius:8px;margin:0 0 16px;font-weight:600}
  .flashmsg.ok{background:#e7f6ec;color:#146c2e;border:1px solid #a7dcb8}
  .flashmsg.bad{background:#fbe9e7;color:#b3261e;border:1px solid #f3c0b8}

  .sitefoot{background:var(--navy);color:#cfe0f5;margin-top:40px;
            display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));
            gap:26px;padding:32px 26px 0}
  .sitefoot .footbrand{font-family:var(--serif);font-size:18px;color:#fff;font-weight:700;margin-bottom:8px}
  .sitefoot .foothead{font-family:var(--mono);font-size:11px;text-transform:uppercase;
                      letter-spacing:.1em;color:var(--navy-soft);margin-bottom:10px}
  .sitefoot p{font-size:13.5px;line-height:1.6;margin:0 0 10px;max-width:34ch}
  .sitefoot a{display:block;color:#cfe0f5;text-decoration:none;font-size:14px;padding:3px 0}
  .sitefoot a:hover{color:#fff;text-decoration:underline}
  .sitefoot .footbar{grid-column:1/-1;border-top:1px solid rgba(255,255,255,.16);
                     margin-top:22px;padding:16px 0;font-size:13px;color:var(--navy-soft);text-align:center}
  .sitefoot .footbar b{color:#fff}
  @media(max-width:640px){.wrap{padding:24px 16px}.page-h1{font-size:26px}}
"""


def nav(active):
    keys = {"home":"", "guide":"", "about":"", "contact":"", "privacy":""}
    if active in keys:
        keys[active] = ' class="active"'
    return NAV.format(**keys)


def content_page(title, active, body):
    """Wrap body content in the full HTML shell with head, nav, footer."""
    return (
        "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\">"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">"
        "<title>%s — RegistrySearch</title>%s<style>%s</style></head><body>%s"
        "<main class=\"wrap\">%s</main>%s</body></html>"
    ) % (title, ADSENSE_HEAD, CONTENT_CSS, nav(active), body, FOOTER)


# ============================================================================
# Content page bodies — real, original content for AdSense + genuine use
# ============================================================================

GUIDE_BODY = r"""
<h1 class="page-h1">How to Search Registry Exports</h1>
<p class="lede">A step-by-step guide to exporting property-registration records
from the Maharashtra IGR portal and searching across many of them at once with
this tool. <span lang="mr">(नोंदणी अभिलेख कसे शोधावेत — मार्गदर्शक)</span></p>
<p class="updated">Practical guide · covers exporting, uploading, filtering and reading results</p>

<h2>1. What this tool does</h2>
<p>The Maharashtra registration portal lets you export search results as a
spreadsheet file. Each export is a snapshot of a single search — one area, one
period, one set of criteria. In real work you end up with many such files, and
answering a simple question ("did this person buy or sell anything across all
of these?") means opening each file by hand.</p>
<p>This tool loads any number of those exports together and searches across all
of them at once. You drop the files in, type what you are looking for, and it
returns matching rows from every file in one table. Nothing is stored on a
server between visits — the files you add are read only when you press Search.</p>
<div class="mr"><span class="lbl">मराठी सारांश</span>
हे साधन अनेक निर्यात (export) केलेल्या एक्सेल फायली एकत्र उघडते आणि सर्व फायलींमध्ये
एकाच वेळी शोध घेते. फायली फक्त "Search" दाबल्यावरच वाचल्या जातात.</div>

<h2>2. Exporting records from the IGR portal</h2>
<p>The public search service of the Department of Registration &amp; Stamps
lets you look up registered documents and download the results. In broad terms:</p>
<ol>
  <li>Open the department's public property-search service and choose the
      search type (for example, search by area, by party name, or by document
      number).</li>
  <li>Select the district, the sub-registrar office (SRO) and the year or date
      range you need.</li>
  <li>Run the search. The portal shows a results grid of matching documents.</li>
  <li>Use the export or download option to save those results as a spreadsheet
      file. This file is what you load into this tool.</li>
</ol>
<div class="callout"><div class="h">A note on the file format</div>
These exports usually arrive with an <code>.xls</code> extension but are, in
fact, HTML tables saved in a Unicode (UTF-16) encoding. That is why Excel often
shows a warning when you open one. This tool reads that format directly, so you
do not need to convert or "fix" the files first — drop them in as they came off
the portal.</div>

<h2>3. Adding files to the tool</h2>
<p>On the search page there is a drop zone on the left. You can:</p>
<ul>
  <li><b>Drag files onto it</b> — one file, several files, or a whole folder.
      Folders are read right through, including sub-folders.</li>
  <li><b>Click it</b> to open the file picker and choose files manually.</li>
</ul>
<p>Each new selection replaces the previous one, so if you load a fresh set of
files the old ones are cleared out and will not turn up in later results.
Accepted types are <code>.xls</code>, <code>.xlsx</code>, <code>.xlsm</code>,
<code>.csv</code>, <code>.htm</code> and <code>.html</code>.</p>

<h2>4. Searching and filtering</h2>
<p>Every filter is optional. Leave them all empty and press Search to see every
row in every file. Fill in one or more to narrow down:</p>
<ul>
  <li><b>Search any column</b> — the broad box at the top. Type any word,
      number or name and it looks across every field of every row.</li>
  <li><b>Document no</b> — the registration document number.</li>
  <li><b>Document type</b> — the kind of deed (sale, gift, release, and so on).</li>
  <li><b>Seller</b> and <b>Purchaser</b> — party names. These match loosely,
      ignoring spaces, quotes and braces, because the portal often exports
      names run together without spaces.</li>
  <li><b>Property contains</b> — any part of the property description: a survey
      number, building name, area, or flat number.</li>
</ul>
<p>Combinations work as an "and" — filling in both a document type and a
purchaser returns rows that match both. Press <b>Enter</b> from any box, or the
Search button, to run it.</p>
<div class="mr"><span class="lbl">मराठी सारांश</span>
प्रत्येक filter ऐच्छिक आहे. विक्रेता व खरेदीदार यांची नावे "loose" पद्धतीने जुळतात —
म्हणजे स्पेस, कंस वगैरे दुर्लक्ष करून. सर्व filter रिकामे ठेवून Search दाबल्यास सर्व
नोंदी दिसतात.</div>

""" + ad_unit() + r"""
<h2>5. Reading the results</h2>
<p>Results appear below in a table, paged so large result sets stay quick to
scan. If you searched the "any column" box, matches are highlighted. Long
fields such as the property description and party names wrap fully so nothing
is hidden. The <b>Download CSV</b> button saves whatever the current filters
return as a spreadsheet you can open in Excel — encoded so Marathi text
displays correctly.</p>

<h2>6. What each column means</h2>
<p>The exports share a common set of columns. The most useful ones:</p>
<table class="fields">
  <tr><th>Column</th><th>What it holds</th></tr>
  <tr><td>docno</td><td>The registration document number for that entry.</td></tr>
  <tr><td>docname</td><td>Type of document — sale deed, gift deed, release, lease, and so on.</td></tr>
  <tr><td>registrationdate</td><td>Date the document was registered at the SRO.</td></tr>
  <tr><td>sroname</td><td>The sub-registrar office where it was registered.</td></tr>
  <tr><td>sellerparty</td><td>Name(s) of the selling / executing party.</td></tr>
  <tr><td>purchaserparty</td><td>Name(s) of the purchasing / claimant party.</td></tr>
  <tr><td>propertydescription</td><td>Free-text description: survey/CTS number, building, area, flat, and extents.</td></tr>
  <tr><td>areaname</td><td>The village or area the property sits in.</td></tr>
</table>

<h2>7. Common pitfalls</h2>
<ul>
  <li><b>Names not matching.</b> The portal often exports party names without
      spaces and inside braces. Use the Seller/Purchaser boxes (which match
      loosely) rather than the exact-text approach, and try a shorter fragment
      of the name.</li>
  <li><b>Expecting the file to be a normal Excel workbook.</b> It is HTML in
      disguise; do not re-save it through another program before loading, as
      that can change the encoding and garble Marathi text.</li>
  <li><b>Very large result sets.</b> A completely empty search over many files
      returns everything; add at least one filter when you only need a slice.</li>
  <li><b>Mixed columns.</b> If some files have slightly different columns, the
      tool still loads them together and simply leaves missing fields blank.</li>
</ul>

<a class="cta" href="/">Open the search tool →</a>
"""

ABOUT_BODY = r"""
<h1 class="page-h1">About IGR &amp; This Tool</h1>
<p class="lede">What the Maharashtra IGR is, why property registration matters,
and why we built a tool to search its exported records.
<span lang="mr">(आय.जी.आर. व या साधनाविषयी)</span></p>
<p class="updated">Background &amp; context</p>

<h2>What is the IGR?</h2>
<p>IGR stands for the <b>Inspector General of Registration</b>, the head of the
Department of Registration &amp; Stamps under the Government of Maharashtra. The
department administers the registration of documents and the collection of
stamp duty across the state, through a network of sub-registrar offices (SROs)
in every district and taluka.</p>
<p>When a property changes hands — a flat is sold, land is gifted, a lease is
granted — the transaction is recorded through this department. The registered
document, and the index of it, becomes part of the public record.</p>
<div class="mr"><span class="lbl">मराठी सारांश</span>
आय.जी.आर. म्हणजे <b>महानिरीक्षक नोंदणी</b> — महाराष्ट्र शासनाचा नोंदणी व मुद्रांक विभाग.
मालमत्तेचे व्यवहार (खरेदी, विक्री, बक्षीस, भाडेपट्टा) याच विभागामार्फत नोंदवले जातात.</div>

<h2>Why property registration is needed</h2>
<p>Registration serves several purposes that matter to anyone dealing with
property:</p>
<ul>
  <li><b>Legal proof of ownership.</b> A registered document is strong evidence
      of the transaction and the parties to it.</li>
  <li><b>A public record.</b> Because registrations are indexed and searchable,
      a prospective buyer or a lawyer can trace the history of a property — who
      owned it, when it changed hands, and on what terms.</li>
  <li><b>Stamp duty and revenue.</b> Registration is the point at which stamp
      duty and registration fees are assessed and paid.</li>
  <li><b>Preventing fraud and disputes.</b> A searchable record of prior
      dealings makes it far harder to sell the same property twice or to hide
      an earlier claim.</li>
</ul>

<h2>How these records are used</h2>
<p>Lawyers, title-search professionals, lenders, and buyers rely on registry
records constantly:</p>
<ul>
  <li><b>Title verification</b> before a purchase or a loan — establishing a
      clean chain of ownership.</li>
  <li><b>Due diligence</b> — checking whether a party has other dealings, or
      whether a property carries earlier transactions.</li>
  <li><b>Litigation support</b> — assembling the documentary history of a
      disputed property.</li>
</ul>

""" + ad_unit() + r"""
<h2>Why we built this tool</h2>
<p>The department's public search service is excellent for one lookup at a
time. But professional work is rarely one lookup. A single title search can
involve dozens of exported result files — several areas, several years, several
name variations — and the real questions cut <em>across</em> all of them:
"Across everything I have pulled, where does this name appear, as buyer or
seller?"</p>
<p>Answering that by opening each spreadsheet by hand is slow and error-prone.
This tool was built to close that gap: load every export together, search and
filter across all of them at once, and get one clean, exportable table of
matches. It reads the portal's exact file format, handles Marathi text
correctly, and keeps nothing on a server between sessions.</p>
<div class="callout"><div class="h">On the data</div>
This tool does not connect to any government system and does not hold a copy of
the registry. It searches the export files that <em>you</em> load, which are
themselves drawn from records the department already publishes. The underlying
registration records are public documents.</div>

<a class="cta" href="/guide">Read the how-to guide →</a>
"""

PRIVACY_BODY = r"""
<h1 class="page-h1">Privacy Policy</h1>
<p class="lede">How this tool handles the files you load and the information you
send through the contact form.</p>
<p class="updated">Please read alongside the notes on the search page.</p>

<h2>Files you load</h2>
<p>When you add Excel or CSV files to the search tool, they are sent to the
server only so they can be read for your search. They are held on temporary
storage for the duration of your session and are cleared when you clear your
selection, when your session ends, or when the server restarts. The files are
not shared, not indexed, and not retained as a permanent database. Each new
selection replaces the previous one.</p>

<h2>What we do not do</h2>
<ul>
  <li>We do not build or keep a central copy of registry data.</li>
  <li>We do not sell or share the contents of files you load.</li>
  <li>We do not require an account or ask you to sign in.</li>
</ul>

<h2>Contacting us</h2>
<p>The contact page lists an email address. If you write to us, we use your
message only to reply to your enquiry.</p>

<h2>Advertising</h2>
<p>Some pages on this site (such as the guide and about pages) may display ads
served by Google, through the Google AdSense programme. Google and its partners
may use cookies to serve ads based on a user's visits to this and other
websites. You can review and control how Google uses data for advertising, and
opt out of personalised advertising, through Google's Ads Settings. The search
tool page itself does not carry advertising.</p>

<h2>Analytics and cookies</h2>
<p>The tool uses a small cookie only to keep your file selection tied to your
own browser session. Advertising on the content pages may set additional
cookies as described above.</p>

<h2>Contact</h2>
<p>Questions about this policy can be sent through the
<a href="/contact">contact page</a>.</p>
"""

def contact_body():
    return r"""
<h1 class="page-h1">Contact</h1>
<p class="lede">Questions about the tool, the guide, or a problem with a file?
Reach us by email and we will reply.</p>
<div class="callout">
  <div class="h">Email</div>
  <p style="margin:0"><a href="mailto:%s" class="emaillink">%s</a></p>
</div>
<p>We usually respond within a couple of working days. When writing about a
file that will not load, it helps to mention which registry office (SRO) and
period the export came from.</p>
""" % (CONTACT_EMAIL, CONTACT_EMAIL)


@app.route("/guide")
def guide():
    return content_page("How-to Guide", "guide", GUIDE_BODY)


@app.route("/about")
def about():
    return content_page("About IGR", "about", ABOUT_BODY)


@app.route("/privacy")
def privacy():
    return content_page("Privacy Policy", "privacy", PRIVACY_BODY)


@app.route("/contact")
def contact():
    return content_page("Contact", "contact", contact_body())



PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>RegistrySearch — Search Maharashtra IGR Property Records</title>
<script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-5876158560001890" crossorigin="anonymous"></script>
<style>
:root{
    --ink:#111827;
    --ink-2:#374151;
    --muted:#6b7280;
    --paper:#f9fafb;
    --sheet:#ffffff;
    --well:#f3f4f6;
    --line:#e5e7eb;
    --line-2:#f3f4f6;
    --navy:#1e3a5f;
    --navy-2:#2563eb;
    --navy-3:#1d4ed8;
    --navy-tint:#e0eaf7;
    --navy-hover:#eff6ff;
    --navy-soft:#93b4d8;
    --gold:#a5540a;
    --gold-tint:#fbe6c0;
    --warn:#b3261e;
    --warn-tint:#fbe9e7;
    --sans:"Segoe UI",Roboto,Inter,system-ui,-apple-system,sans-serif;
    --serif:Georgia,"Times New Roman",serif;
    --mono:ui-monospace,"SF Mono",Consolas,"Liberation Mono",monospace;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--paper);color:var(--ink);font-family:var(--sans);
       font-size:15px;line-height:1.55;-webkit-font-smoothing:antialiased;
       font-variant-numeric:tabular-nums}

  .app{display:grid;grid-template-columns:minmax(0,1fr);gap:20px;
       min-height:100vh;padding:0 0 32px;max-width:1720px;margin:0 auto}

  .sheet{background:var(--sheet);border:1px solid var(--line);
         border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,0.06);
         display:flex;flex-direction:column;min-width:0;overflow:hidden;
         margin:0 24px}
  .masthead{border-radius:0;margin:0}

  /* ---------- masthead ---------- */
  .masthead{background:#1e3a5f;color:#fff;padding:18px 28px;
            display:flex;align-items:center;gap:14px;flex-wrap:wrap;
            border-bottom:3px solid #2563eb;margin-bottom:20px}
  .masthead .firm{font-family:var(--serif);font-size:26px;font-weight:700;
                  color:#fff;letter-spacing:.01em;line-height:1.1}
  .masthead .firm .amp{color:#93b4d8;font-style:italic;padding:0 2px}
  .masthead .tool{font-family:var(--mono);font-size:11px;text-transform:uppercase;
                  letter-spacing:.22em;color:#e0eaf7;font-weight:600;
                  align-self:center;padding-left:14px;
                  border-left:1px solid rgba(255,255,255,.25)}
  .mastnav{display:flex;gap:2px;flex-wrap:wrap}
  .mastnav a{color:#cfe0f5;text-decoration:none;font-size:13px;font-weight:600;
             padding:6px 12px;border-radius:6px}
  .mastnav a:hover{background:rgba(255,255,255,.1);color:#fff}
  .mastnav a.active{background:rgba(255,255,255,.16);color:#fff}
  .masthead .by{margin-left:auto;font-size:13px;color:#93b4d8}
  .masthead .by b{font-family:var(--serif);font-style:italic;font-weight:700;
                  color:#fff;font-size:15px}

  /* ---------- band 1 ---------- */
  .qgrid{display:grid;grid-template-columns:300px minmax(0,1fr)}
  .src{padding:20px;background:#1e3a5f;border-right:1px solid rgba(255,255,255,.22);border-radius:10px 0 0 0}
  .filters{padding:20px 24px 24px;min-width:0;background:#1e3a5f;border-radius:0 10px 0 0}

  .cap{font-family:var(--sans);font-size:11px;text-transform:uppercase;
       letter-spacing:.12em;color:#fff;font-weight:700;
       display:flex;align-items:center;gap:10px;margin-bottom:14px}
  .cap .n{margin-left:auto;color:#fff;letter-spacing:.02em;
          background:rgba(255,255,255,.14);padding:2px 9px;border:1px solid rgba(255,255,255,.3);
          font-size:11px;border-radius:10px;font-weight:700}
  .cap .hint{text-transform:none;letter-spacing:0;font-family:var(--sans);
             font-size:12px;color:#e0eaf7;font-weight:600}

  .drop{border:2px dashed rgba(255,255,255,.35);background:rgba(255,255,255,.06);padding:24px 16px;
        text-align:center;cursor:pointer;display:block;width:100%;font:inherit;
        color:#fff;transition:border-color .12s,background .12s;border-radius:8px}
  .drop:hover{border-color:#7ab0f0;background:rgba(255,255,255,.12)}
  .drop.over{border-color:#7ab0f0;background:rgba(255,255,255,.14);border-style:solid}
  .drop .big{font-weight:700;font-size:14px;display:block;color:#fff}
  .drop .small{font-size:12px;color:#e0eaf7;display:block;margin-top:4px;font-weight:600}
  .drop svg{opacity:.7;margin-bottom:8px;color:#bcd3ee}

  .chips{margin-top:14px;max-height:150px;overflow:auto;font-size:12.5px;color:#fff;font-weight:600}
  .chip{padding:5px 0;border-bottom:1px solid rgba(255,255,255,.22);display:flex;gap:8px}
  .chip span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .chip em{margin-left:auto;color:#e0eaf7;font-style:normal;flex:none}

  label{display:block;font-size:11px;font-family:var(--sans);text-transform:uppercase;
        letter-spacing:.08em;color:#fff;margin-bottom:5px;font-weight:700}
  input[type=text]{
    width:100%;padding:9px 12px;border:1.5px solid #3a577a;
    background:#fff;font-family:inherit;font-size:14px;color:var(--ink);
    border-radius:6px;outline:none;transition:border-color .15s,box-shadow .15s}
  input[type=text]:hover{border-color:#6f93bd}
  input[type=text]:focus{border-color:#7ab0f0;box-shadow:0 0 0 3px rgba(122,176,240,0.35)}
  input::placeholder{color:#9ca3af}
  .drop:focus-visible,button:focus-visible{outline:2.5px solid #2563eb;outline-offset:2px}
  .field{margin-bottom:14px}
  .fgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:0 18px}
  .any{margin:0 0 18px;padding:0 0 18px;background:none;
       border:0;border-bottom:1px solid rgba(255,255,255,.3);border-radius:0}
  .any input{font-size:15px;padding:10px 12px;border-color:#93b4d8}
  .any input:focus{border-color:#7ab0f0;box-shadow:0 0 0 3px rgba(122,176,240,0.35)}
  .any label{font-size:12px;color:#fff;font-weight:700}

  button{font-family:var(--sans);font-size:14px;font-weight:700;padding:10px 22px;
         border:none;background:#2563eb;color:#fff;cursor:pointer;
         border-radius:7px;letter-spacing:.02em;transition:.12s}
  button:hover:not(:disabled){background:#1d4ed8}
  button.ghost{background:#fff;color:#1e3a5f;border:1.5px solid #2563eb}
  button.ghost:hover:not(:disabled){background:var(--navy-hover)}
  button.link{background:none;border:0;color:#bcd3ee;padding:0;font-size:12px;
              font-family:var(--sans);text-decoration:underline;text-underline-offset:3px;
              font-weight:600;letter-spacing:0}
  button.link:hover:not(:disabled){color:#fff;background:none}
  button:disabled{opacity:.45;cursor:not-allowed}

  .actions{display:flex;gap:12px;align-items:center;padding:16px 24px;
           border-top:1px solid rgba(255,255,255,.22);background:#1e3a5f;border-radius:0 0 10px 10px}
  .actions .msg{font-size:13px;color:#dbe6f5;margin-left:auto;text-align:right;font-weight:600}

  /* ---------- results ---------- */
  .bar{display:flex;align-items:center;gap:16px;padding:14px 22px;
       border-bottom:2px solid #1e3a5f;flex-wrap:wrap;background:var(--well);
       border-radius:10px 10px 0 0}
  .bar h2{margin:0;font-size:11px;text-transform:uppercase;
          letter-spacing:.14em;color:#2563eb;font-weight:700}
  .tally{font-size:20px;font-weight:700;letter-spacing:-.01em;color:var(--ink)}
  .tally span{color:var(--muted);font-weight:400;font-size:14px}
  .stat{font-family:var(--mono);font-size:12px;color:var(--ink-2);margin-left:auto}
  .flash{padding:11px 22px;background:var(--warn-tint);color:var(--warn);
         font-size:13px;border-bottom:1px solid var(--line);font-weight:500}

  .tablewrap{min-width:0;overflow-x:auto}
  table{border-collapse:collapse;width:100%;font-size:13.5px;table-layout:fixed}
  col.c-docno{width:7%}
  col.c-registrationdate{width:8%}
  col.c-docname{width:9%}
  col.c-areaname{width:8%}
  col.c-sroname{width:11%}
  col.c-sellerparty{width:15%}
  col.c-purchaserparty{width:15%}
  col.c-propertydescription{width:27%}
  th{background:#1e3a5f;text-align:left;padding:11px 14px;
     font-size:11px;text-transform:uppercase;letter-spacing:.06em;
     white-space:nowrap;color:#e0eaf7;font-weight:600;border-bottom:2px solid #2563eb}
  td{padding:11px 14px;border-bottom:1px solid var(--line);vertical-align:top;
     color:var(--ink);line-height:1.5;overflow-wrap:break-word;word-break:break-word}
  tbody tr:nth-child(even) td{background:var(--well)}
  tbody tr:hover td{background:var(--navy-hover)}
  td.party,td.prop{font-size:12.5px;line-height:1.45}
  mark{background:var(--gold-tint);color:var(--ink);padding:0 2px;font-weight:600;border-radius:2px}

  .pager{display:flex;gap:12px;align-items:center;padding:12px 22px;
         border-top:1px solid var(--line);font-size:13px;color:var(--ink-2);
         border-radius:0 0 10px 10px}
  .pager button{padding:7px 14px;font-size:13px}
  .pager select{padding:5px 8px;font-size:13px;border:1.5px solid var(--line);
                background:#fff;color:var(--ink);border-radius:6px;outline:none}
  .pager b{color:var(--ink)}

  .blank{flex:1;min-height:260px;display:grid;place-content:center;text-align:center;
         padding:48px;color:var(--ink-2);gap:8px;max-width:460px;margin:0 auto}
  .blank .t{font-size:17px;font-weight:700;color:var(--ink)}
  .blank .s{font-size:14px;line-height:1.65;color:var(--muted)}
  .blank .glyph{font-size:30px;color:var(--line);margin-bottom:4px}
  .err{color:var(--warn);font-weight:600}
  .hide{display:none!important}
  .spin{display:inline-block;width:11px;height:11px;border:2px solid var(--line);
        border-top-color:#2563eb;border-radius:50%;animation:sp .6s linear infinite;
        vertical-align:-1px;margin-right:7px}
  @keyframes sp{to{transform:rotate(360deg)}}

  @media(max-width:920px){
    .app{padding:0 12px 16px}
    .qgrid{grid-template-columns:1fr}
    .src{border-right:0;border-bottom:1px solid var(--line)}
    .sheet{margin:0 8px}
  }

  /* ---- info content below the tool (same URL as the tool) ---- */
  .infoblock{max-width:900px;margin:6px auto 0;padding:8px 26px 20px;line-height:1.7}
  .ib-h1{font-family:var(--serif);font-size:26px;color:var(--navy);margin:14px 0 6px;line-height:1.25}
  .ib-lede{font-size:16px;color:var(--ink-2);margin:0 0 8px}
  .infoblock h2{font-family:var(--serif);font-size:21px;color:var(--navy);
                margin:28px 0 8px;padding-bottom:6px;border-bottom:2px solid var(--line)}
  .infoblock h3{font-size:16px;color:var(--ink);margin:18px 0 4px}
  .infoblock p{margin:0 0 12px;color:var(--ink)}
  .infoblock ol,.infoblock ul{margin:0 0 14px;padding-left:22px}
  .infoblock li{margin:0 0 7px}
  .infoblock a{color:var(--navy-2);text-decoration:underline;text-underline-offset:2px}
  .infoblock code{background:var(--well);padding:1px 5px;border-radius:3px;font-size:13px}
  .ib-mr{font-size:14.5px;color:var(--ink-2);background:var(--well);
         border-left:3px solid var(--navy-soft);padding:10px 14px;margin:0 0 14px;border-radius:0 4px 4px 0}
  .ib-lbl{font-family:var(--mono);font-size:10px;text-transform:uppercase;letter-spacing:.1em;
          color:var(--muted);display:block;margin-bottom:3px}
  .ib-fields{width:100%;border-collapse:collapse;margin:4px 0 18px;font-size:14.5px}
  .ib-fields th{background:var(--navy);color:#fff;text-align:left;padding:8px 12px;
                font-size:11px;text-transform:uppercase;letter-spacing:.05em}
  .ib-fields td{padding:8px 12px;border-bottom:1px solid var(--line);vertical-align:top}
  .ib-fields tr:nth-child(even) td{background:var(--well)}
  .ib-fields td:first-child{font-family:var(--mono);font-size:12.5px;color:var(--navy);font-weight:600}
  .adwrap{margin:22px 0;text-align:center;min-height:100px;background:var(--well);
          border:1px solid var(--line);border-radius:8px;padding:10px;overflow:hidden}
  .adtag{display:block;font-family:var(--mono);font-size:10px;text-transform:uppercase;
         letter-spacing:.1em;color:var(--muted);margin-bottom:6px;text-align:left}
  .sitefoot{background:var(--navy);color:#cfe0f5;margin-top:36px;
            display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));
            gap:26px;padding:32px 26px 0}
  .sitefoot .footbrand{font-family:var(--serif);font-size:18px;color:#fff;font-weight:700;margin-bottom:8px}
  .sitefoot .foothead{font-family:var(--mono);font-size:11px;text-transform:uppercase;
                      letter-spacing:.1em;color:var(--navy-soft);margin-bottom:10px}
  .sitefoot p{font-size:13.5px;line-height:1.6;margin:0 0 10px;max-width:34ch}
  .sitefoot a{display:block;color:#cfe0f5;text-decoration:none;font-size:14px;padding:3px 0}
  .sitefoot a:hover{color:#fff;text-decoration:underline}
  .sitefoot .footbar{grid-column:1/-1;border-top:1px solid rgba(255,255,255,.16);
                     margin-top:22px;padding:16px 0;font-size:13px;color:var(--navy-soft);text-align:center}
  .sitefoot .footbar b{color:#fff}
</style>
</head>
<body>
<div class="app">

  <header class="masthead">
    <span class="firm">Registry<span class="amp">Search</span></span>
    <span class="tool">Registry Search</span>
    <nav class="mastnav">
      <a href="/" class="active">Search</a>
      <a href="/guide">Guide</a>
      <a href="/about">About IGR</a>
      <a href="/contact">Contact</a>
    </nav>
    <span class="by">Search Maharashtra IGR records</span>
  </header>

  <!-- band 1 — files and filters -->
  <section class="sheet">
    <div class="qgrid">

      <div class="src">
        <div class="cap">Files <span class="n" id="filecount"></span></div>

        <button type="button" class="drop" id="drop">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor"
               stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
            <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
          </svg>
          <span class="big">Drop files or folders</span>
          <span class="small">or click to pick files</span>
        </button>
        <input type="file" id="picker" class="hide" multiple
               accept=".xls,.xlsx,.xlsm,.csv,.htm,.html">

        <div class="chips" id="chips"></div>
        <div style="margin-top:12px">
          <button type="button" class="link hide" id="clearfiles">Clear all files</button>
        </div>
      </div>

      <div class="filters">
        <div class="cap">Filters
          <span class="hint">every box is optional</span>
          <button type="button" class="link n" id="clearfilters">Reset</button>
        </div>

        <div class="any">
          <div class="field" style="margin:0">
            <label for="_all">Search any column</label>
            <input type="text" id="_all" placeholder="type a word, number or name">
          </div>
        </div>

        <div class="fgrid">
          {% for f in filters %}
            <div class="field">
              <label for="f_{{ f.col }}">{{ f.label }}</label>
              <input type="text" id="f_{{ f.col }}" data-col="{{ f.col }}"></div>
          {% endfor %}
        </div>
      </div>
    </div>

    <div class="actions">
      <button id="searchbtn" disabled>Search</button>
      <button class="ghost" id="exportbtn" type="button" disabled>Download CSV</button>
      <span class="msg" id="selmsg"></span>
    </div>
  </section>

  <!-- band 2 — results -->
  <section class="sheet results">
    <div class="bar">
      <h2>Results</h2>
      <span class="tally" id="tally">—</span>
      <span class="stat" id="stat"></span>
    </div>
    <div class="flash hide" id="flash"></div>

    <div class="blank" id="blank">
      <div class="glyph">§</div>
      <div class="t" id="blanktitle">Add files to begin</div>
      <div class="s" id="blanksub">Drop the registry exports above. Nothing is read until you press Search, and leaving every filter empty returns every row.</div>
    </div>

    <div class="tablewrap hide" id="tablewrap">
      <table><thead id="thead"></thead><tbody id="tbody"></tbody></table>
    </div>

    <div class="pager hide" id="pager">
      <button class="ghost" id="prev" type="button">‹ Previous</button>
      <span id="pageinfo"></span>
      <button class="ghost" id="next" type="button">Next ›</button>
      <span style="margin-left:auto">Rows per page
        <select id="size"><option>25</option><option selected>50</option>
          <option>100</option><option>250</option></select></span>
    </div>
  </section>

  <!-- content below the tool: on the SAME url as the tool, so this page
       carries real text of its own (and ad slots sit beside that text) -->
  <section class="infoblock">
    <h1 class="ib-h1">Search Maharashtra Property-Registration (IGR) Records</h1>
    <p class="ib-lede">Load the Excel exports from the Maharashtra registration
      portal and search across all of them at once — by document number, type,
      party name, or property description. This page explains what the tool
      does, how to use it, and how to read what it returns.</p>

    <h2>What this tool is for</h2>
    <p>The Department of Registration &amp; Stamps (IGR), Government of
      Maharashtra, publishes an index of registered documents — sales, gifts,
      releases, leases and more — through its public search service. That
      service returns one search at a time as a downloadable spreadsheet. In
      real title-search or due-diligence work you gather many such exports
      across different areas, years and name spellings, and the questions that
      matter cut across all of them at once: <em>where does this name appear as
      a buyer or a seller, in any of these files?</em></p>
    <p>Opening each spreadsheet by hand to answer that is slow. This tool loads
      every export together and searches them as one set, returning matching
      rows from all of them in a single table you can page through and download.
      Files are read only when you press Search, and nothing is kept on the
      server between sessions.</p>
    <div class="ib-mr"><span class="ib-lbl">मराठी सारांश</span>
      हे साधन महाराष्ट्र नोंदणी विभागाच्या (IGR) निर्यात केलेल्या अनेक एक्सेल फायली
      एकत्र उघडते आणि दस्त क्रमांक, दस्त प्रकार, पक्षकाराचे नाव किंवा मालमत्ता वर्णनानुसार
      सर्व फायलींमध्ये एकाच वेळी शोध घेते.</div>

    <div class="adwrap"><span class="adtag">Advertisement</span>
      <ins class="adsbygoogle" style="display:block" data-ad-client="ca-pub-5876158560001890"
        data-ad-format="auto" data-full-width-responsive="true"></ins>
      <script>(adsbygoogle=window.adsbygoogle||[]).push({});</script></div>

    <h2>How to use it, step by step</h2>
    <ol>
      <li><b>Export from the IGR portal.</b> Run your search on the department's
        public property-search service and download the result as a spreadsheet.
        The file usually has an <code>.xls</code> name but is really an HTML
        table in UTF-16 — this tool reads that directly, so you do not need to
        convert it first.</li>
      <li><b>Add the files.</b> Drag files or a whole folder onto the drop zone
        at the top of this page, or click it to pick files. Each new selection
        replaces the last, so old files never leak into new results.</li>
      <li><b>Filter.</b> Every box is optional. Use <b>Search any column</b> for
        a broad look, or the specific boxes — Document no, Document type, Seller,
        Purchaser, Property contains. Seller and Purchaser match loosely,
        ignoring spaces and braces, because the portal exports names run
        together.</li>
      <li><b>Read and export.</b> Matches appear in the table above, paged for
        speed, with any-column matches highlighted. Download CSV saves the
        current filtered set with Marathi text intact.</li>
    </ol>

    <h2>What the columns mean</h2>
    <table class="ib-fields">
      <tr><th>Column</th><th>What it holds</th></tr>
      <tr><td>docno</td><td>Registration document number.</td></tr>
      <tr><td>docname</td><td>Document type — sale, gift, release, lease, etc.</td></tr>
      <tr><td>registrationdate</td><td>Date registered at the sub-registrar office.</td></tr>
      <tr><td>sroname</td><td>The sub-registrar office (SRO).</td></tr>
      <tr><td>sellerparty / purchaserparty</td><td>Names of the executing and claiming parties.</td></tr>
      <tr><td>propertydescription</td><td>Survey/CTS number, building, area, flat and extents.</td></tr>
    </table>

    <h2>Common questions</h2>
    <h3>Why does Excel warn when opening the export?</h3>
    <p>Because the portal's <code>.xls</code> files are HTML tables in disguise,
      saved in Unicode. That is normal. Load them into this tool as they came
      off the portal — do not re-save them through another program, as that can
      garble the Marathi text.</p>
    <h3>A party name will not match — why?</h3>
    <p>Names are often exported without spaces and inside braces. Use the Seller
      or Purchaser boxes, which match loosely, and try a shorter fragment of the
      name rather than the full string.</p>
    <h3>Is my data uploaded permanently?</h3>
    <p>No. Files are read for your search and cleared when you clear your
      selection or your session ends. See the
      <a href="/privacy">privacy policy</a> for detail.</p>

    <p style="margin-top:22px">More detail is in the
      <a href="/guide">full how-to guide</a> and the
      <a href="/about">background on IGR and this tool</a>.</p>
  </section>
</div>

<footer class="sitefoot">
  <div class="footcol">
    <div class="footbrand">RegistrySearch</div>
    <p>A search tool for Maharashtra property-registration (IGR) records, built
      to look across many registry exports at once.</p>
  </div>
  <div class="footcol">
    <div class="foothead">Pages</div>
    <a href="/">Registry Search</a>
    <a href="/guide">How-to Guide</a>
    <a href="/about">About IGR &amp; This Tool</a>
    <a href="/contact">Contact</a>
    <a href="/privacy">Privacy Policy</a>
  </div>
  <div class="footcol">
    <div class="foothead">Note</div>
    <p>IGR registration records are public documents published by the Department
      of Registration &amp; Stamps, Government of Maharashtra. This tool searches
      exports of that public data.</p>
  </div>
  <div class="footbar">RegistrySearch · search Maharashtra IGR records</div>
</footer>

<script>
const $ = id => document.getElementById(id);
let ROWS = [], COLS = [], WIDE = [], PAGE = 1, NFILES = 0, HILITE = "";
const OK_EXT = /\.(xls|xlsx|xlsm|csv|html?)$/i;
// columns rendered full-width with wrapping (long Marathi text)
const PARTY = new Set(["sellerparty","purchaserparty"]);   // smaller type
const PROP  = new Set(["propertydescription"]);

const esc = s => String(s ?? "").replace(/[&<>"']/g, c =>
  ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

// ---------- sources ----------
$("drop").onclick = () => $("picker").click();
$("picker").onchange = e => send([...e.target.files]);

["dragenter","dragover"].forEach(ev => $("drop").addEventListener(ev, e => {
  e.preventDefault(); $("drop").classList.add("over");
}));
["dragleave","drop"].forEach(ev => $("drop").addEventListener(ev, e => {
  e.preventDefault(); $("drop").classList.remove("over");
}));

$("drop").addEventListener("drop", async e => {
  const items = [...(e.dataTransfer.items || [])];
  const entries = items.map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()).filter(Boolean);
  if (entries.length) {
    setMsg("reading folder…", false, true);
    const files = [];
    for (const en of entries) await walk(en, files);
    send(files);
  } else {
    send([...e.dataTransfer.files]);
  }
});

function walk(entry, out){
  return new Promise(res => {
    if (entry.isFile) return entry.file(f => { out.push(f); res(); }, res);
    const rd = entry.createReader();
    const readMore = () => rd.readEntries(async batch => {
      if (!batch.length) return res();
      for (const e of batch) await walk(e, out);
      readMore();
    }, res);
    readMore();
  });
}

function setMsg(t, bad, busy){
  $("selmsg").innerHTML = (busy ? '<span class="spin"></span>' : "") + esc(t);
  $("selmsg").className = bad ? "msg err" : "msg";
}

// a new upload REPLACES the old set; wipe any results still on screen
function clearResults(){
  ROWS = []; COLS = []; PAGE = 1;
  $("tablewrap").classList.add("hide"); $("pager").classList.add("hide");
  $("flash").classList.add("hide");
  $("tally").textContent = "—"; $("stat").textContent = "";
  $("thead").innerHTML = ""; $("tbody").innerHTML = "";
  $("blank").classList.remove("hide");
}

async function send(files){
  const usable = files.filter(f => OK_EXT.test(f.name) && !f.name.startsWith("~$"));
  if (!usable.length) {
    if (files.length) setMsg("no Excel or CSV files in that selection", true);
    return;
  }
  const fd = new FormData();
  usable.forEach(f => fd.append("files", f));

  setMsg("sending " + usable.length + " file(s)…", false, true);
  try {
    const r = await fetch("/api/select", {method:"POST", body:fd});
    const j = await r.json();
    if (!r.ok) throw new Error(j.error || "could not add those");
    NFILES = j.files;
    clearResults();                       // <- old results gone with the old files
    $("filecount").textContent = j.files + (j.files === 1 ? " file" : " files");
    $("chips").innerHTML = j.names.map(n => `<div class="chip"><span>${esc(n)}</span></div>`).join("") +
      (j.files > j.names.length ? `<div class="chip"><em>…and ${j.files - j.names.length} more</em></div>` : "");
    $("clearfiles").classList.remove("hide");
    setMsg("not read yet — press Search", false);
    $("searchbtn").disabled = false; $("exportbtn").disabled = false;
    $("blanktitle").textContent = "Ready to search";
    $("blanksub").textContent = "Press Search. Leaving every filter empty returns every row.";
    $("picker").value = "";
  } catch (e) { setMsg(e.message, true); }
}

$("clearfiles").onclick = async () => {
  await fetch("/api/clear", {method:"POST"});
  NFILES = 0;
  clearResults();
  $("chips").innerHTML = ""; $("filecount").textContent = "";
  $("clearfiles").classList.add("hide");
  $("searchbtn").disabled = true; $("exportbtn").disabled = true;
  $("blanktitle").textContent = "Add files to begin";
  $("blanksub").textContent = "Drop the registry exports above. Nothing is read until you press Search.";
  setMsg("cleared — uploaded copies deleted", false);
};

// ---------- filters ----------
function collect(){
  const p = {};
  const free = $("_all").value.trim();
  if (free) p._all = free;
  document.querySelectorAll(".fgrid [data-col]").forEach(el => {
    if (el.value.trim()) p[el.dataset.col] = el.value.trim();
  });
  return p;
}

$("clearfilters").onclick = () => {
  $("_all").value = "";
  document.querySelectorAll(".fgrid [data-col]").forEach(el => el.value = "");
};

// ---------- search ----------
$("searchbtn").onclick = doSearch;
$("size").onchange = () => showPage(1);
$("prev").onclick = () => showPage(PAGE - 1);
$("next").onclick = () => showPage(PAGE + 1);
document.addEventListener("keydown", e => {
  if (e.key === "Enter" && !$("searchbtn").disabled) doSearch();
});

async function doSearch(){
  $("searchbtn").disabled = true;
  $("stat").innerHTML = '<span class="spin"></span>opening ' + NFILES + ' file(s)…';
  $("flash").classList.add("hide");
  const t0 = performance.now();
  try {
    const r = await fetch("/api/search", {method:"POST",
      headers:{"Content-Type":"application/json"}, body: JSON.stringify(collect())});
    const j = await r.json();
    if (!r.ok) throw new Error(j.error || "search failed");

    ROWS = j.rows; COLS = j.columns; WIDE = j.long || []; HILITE = ($("_all").value || "").trim();
    $("tally").innerHTML = "<b>" + j.total.toLocaleString() + "</b> <span>" +
      (j.total === 1 ? "row" : "rows") + "</span>";
    $("stat").textContent = j.files + " files · " + j.scanned.toLocaleString() +
      " scanned · " + ((performance.now() - t0) / 1000).toFixed(1) + "s";

    const bad = (j.notes || []).filter(n => !n.ok);
    const msgs = [];
    if (j.truncated) msgs.push("Stopped at the first " + j.cap.toLocaleString() +
      " matches to keep server memory low — narrow the filters to see the rest.");
    if (bad.length) msgs.push(bad.length + " file(s) unreadable: " +
      bad.map(n => esc(n.file)).join(", "));
    $("flash").classList.toggle("hide", !msgs.length);
    $("flash").innerHTML = msgs.join(" ");

    showPage(1);
  } catch (e) {
    $("stat").innerHTML = '<span class="err">' + esc(e.message) + '</span>';
  }
  $("searchbtn").disabled = false;
}

function mark(html){
  if (!HILITE) return html;
  try {
    const re = new RegExp("(" + HILITE.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "gi");
    return html.replace(re, "<mark>$1</mark>");
  } catch { return html; }
}

function showPage(page){
  const size = parseInt($("size").value, 10);
  const pages = Math.max(1, Math.ceil(ROWS.length / size));
  PAGE = Math.min(Math.max(1, page), pages);

  const has = ROWS.length > 0;
  $("tablewrap").classList.toggle("hide", !has);
  $("pager").classList.toggle("hide", !has);
  $("blank").classList.toggle("hide", has);
  if (!has) {
    $("blanktitle").textContent = "No rows matched";
    $("blanksub").textContent = "Clear or widen a filter and search again.";
    return;
  }

  $("thead").innerHTML =
    "<colgroup>" + COLS.map(c => `<col class="c-${esc(c)}">`).join("") + "</colgroup>" +
    "<tr>" + COLS.map(c => `<th>${esc(c.replace(/_/g," "))}</th>`).join("") + "</tr>";

  // full text shown outright — no clamp, no "more" toggle
  $("tbody").innerHTML = ROWS.slice((PAGE-1)*size, PAGE*size).map(row =>
    "<tr>" + row.map((val, i) => {
      const c = COLS[i], v = mark(esc(val));
      const cls = PARTY.has(c) ? "party" : PROP.has(c) ? "prop" : "";
      return `<td class="${cls}">${v}</td>`;
    }).join("") + "</tr>").join("");

  const from = (PAGE-1)*size + 1, to = Math.min(PAGE*size, ROWS.length);
  $("pageinfo").innerHTML = `<b>${from}–${to}</b> of <b>${ROWS.length.toLocaleString()}</b>`;
  $("prev").disabled = PAGE <= 1;
  $("next").disabled = to >= ROWS.length;
  window.scrollTo({top: $("tablewrap").getBoundingClientRect().top + window.scrollY - 12, behavior:"smooth"});
}

// ---------- export ----------
$("exportbtn").onclick = async () => {
  $("exportbtn").disabled = true;
  const was = $("stat").innerHTML;
  $("stat").innerHTML = '<span class="spin"></span>building CSV…';
  try {
    const r = await fetch("/api/export", {method:"POST",
      headers:{"Content-Type":"application/json"}, body: JSON.stringify(collect())});
    if (!r.ok) throw new Error("export failed");
    const blob = await r.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "results.csv"; a.click();
    URL.revokeObjectURL(a.href);
    $("stat").innerHTML = was;
  } catch (e) { $("stat").innerHTML = '<span class="err">' + esc(e.message) + '</span>'; }
  $("exportbtn").disabled = false;
};
</script>
</body>
</html>
"""


def lan_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


if __name__ == "__main__":
    shutil.rmtree(UPLOAD_DIR, ignore_errors=True)
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    ip = lan_ip()
    print("\n  Registry Search is running.")
    print(f"  This PC:      http://127.0.0.1:{PORT}")
    print(f"  Other PCs:    http://{ip}:{PORT}")
    print(f"  Upload spool: {UPLOAD_DIR}\n")
    try:
        from waitress import serve
        serve(app, host=HOST, port=PORT, threads=4)
    except ImportError:
        app.run(host=HOST, port=PORT, threaded=True, debug=False)
