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
import sys
import uuid
import codecs
import socket
import shutil
import tempfile
import hmac
import threading
from datetime import datetime, timedelta

from lxml import etree
from flask import (Flask, Request, request, jsonify, render_template_string,
                   Response, make_response, stream_with_context, session,
                   redirect, url_for)

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

# Leave empty for no login (fine on a LAN). Set the APP_PASSWORD environment
# variable on the host to require one shared password, entered once per
# browser and remembered for 30 days.
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

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

# Columns offered as filters, in the order they appear in the UI.
#   type "text"   -> case-insensitive "contains" box
#   type "number" -> min / max pair
#   type "date"   -> from / to pair
#   loose: True   -> ignores spaces, quotes and braces when matching (good for
#                    Marathi party names that are exported without spaces)
FILTERS = [
    {"col": "docno",                  "label": "Document no",        "type": "text"},
    {"col": "internaldocumentnumber", "label": "Internal doc no",    "type": "text"},
    {"col": "sellerparty",            "label": "Seller",             "type": "text", "loose": True},
    {"col": "purchaserparty",         "label": "Purchaser",          "type": "text", "loose": True},
    {"col": "propertydescription",    "label": "Property contains",  "type": "text", "loose": True},
    {"col": "areaname",               "label": "Area",               "type": "text"},
    {"col": "docname",                "label": "Document type",      "type": "text"},
    {"col": "sroname",                "label": "SRO",                "type": "text"},
    {"col": "srocode",                "label": "SRO code",           "type": "text"},
    {"col": "status",                 "label": "Status",             "type": "text"},
    {"col": "registrationdate",       "label": "Registration date",  "type": "date"},
    {"col": "consideration_amt",      "label": "Consideration",      "type": "number"},
    {"col": "marketvalue",            "label": "Market value",       "type": "number"},
]

# Columns shown in the results table, in order. Leave empty to show everything.
DISPLAY_COLUMNS = [
    "docno", "registrationdate", "docname", "areaname", "sroname",
    "sellerparty", "purchaserparty", "propertydescription",
    "consideration_amt", "marketvalue", "stampdutypaid", "registrationfees",
]

# Columns rendered as long text with a "more" toggle instead of one line.
LONG_TEXT_COLUMNS = ["propertydescription", "sellerparty", "purchaserparty"]

# Formats tried when reading a date column, first match wins.
DATE_FORMATS = ["%d/%m/%Y", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y", "%Y-%m-%d"]

FILE_EXTENSIONS = (".xls", ".xlsx", ".xlsm", ".csv", ".htm", ".html")

# ============================================================================
# Small helpers
# ============================================================================

_LOOSE_RE = re.compile(r'[\s"\'{}\[\]().,\-_/]')
_NUM_RE = re.compile(r"[^0-9.\-]")


def loose(s):
    return _LOOSE_RE.sub("", s).lower()


def to_number(s):
    try:
        return float(_NUM_RE.sub("", str(s)))
    except (ValueError, TypeError):
        return None


def to_date(s):
    s = s.strip()
    if not s:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


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

        elif f["type"] == "number":
            lo = to_number(params.get(col + "__min", "") or "")
            hi = to_number(params.get(col + "__max", "") or "")
            if lo is None and hi is None:
                continue

            def num_test(row, i=i, lo=lo, hi=hi):
                v = to_number(cell(row, i))
                if v is None:
                    return False
                return (lo is None or v >= lo) and (hi is None or v <= hi)
            preds.append(num_test)

        elif f["type"] == "date":
            raw_lo = params.get(col + "__from", "")
            raw_hi = params.get(col + "__to", "")
            lo = datetime.strptime(raw_lo, "%Y-%m-%d").date() if raw_lo else None
            hi = datetime.strptime(raw_hi, "%Y-%m-%d").date() if raw_hi else None
            if lo is None and hi is None:
                continue

            def date_test(row, i=i, lo=lo, hi=hi):
                d = to_date(cell(row, i))
                if d is None:
                    return False
                return (lo is None or d >= lo) and (hi is None or d <= hi)
            preds.append(date_test)

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
                    matches.append([cell(raw, i) if i >= 0 else "" for i in take] + [name])

            notes.append({"file": name, "rows": rows_here, "ok": True})
        except Exception as exc:
            notes.append({"file": name, "rows": 0, "ok": False,
                          "error": f"{type(exc).__name__}: {exc}"})
        finally:
            it = None

        if truncated:
            break

    return {
        "columns": (columns or []) + ["_source_file"],
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
app.secret_key = os.environ.get("SECRET_KEY") or os.urandom(32)
app.permanent_session_lifetime = timedelta(days=30)
# Marathi text as raw UTF-8 rather than \uXXXX escapes — roughly halves the
# size of every response, in RAM and on the wire.
app.json.ensure_ascii = False
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

# Behind a platform's HTTPS proxy, trust the forwarded scheme so cookies are
# marked Secure correctly.
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


def session_dir(sid):
    d = os.path.join(UPLOAD_DIR, sid)
    os.makedirs(d, exist_ok=True)
    return d


def drop_session(sid):
    s = SESSIONS.pop(sid, None)
    if s and s.get("dir"):
        shutil.rmtree(s["dir"], ignore_errors=True)


def add_to_session(sid, new_paths, updir):
    """Selections accumulate: drop a folder, then add loose files, and all of
    them are searched together."""
    with _lock:
        s = SESSIONS.get(sid) or {"paths": [], "dir": updir}
        s["paths"] = list(s.get("paths") or []) + list(new_paths)
        s["dir"] = updir
        s["touched"] = datetime.now()
        SESSIONS[sid] = s
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


@app.route("/healthz")
def healthz():
    """Cheap endpoint for an uptime pinger. Keeping this warm every 10 minutes
    stops a free-tier host from sleeping and wiping everyone's uploads."""
    return "ok", 200


@app.before_request
def require_password():
    """No-op unless APP_PASSWORD is set on the host."""
    if not APP_PASSWORD:
        return None
    if request.endpoint in ("login", "healthz") or request.path == "/login":
        return None
    if session.get("ok"):
        return None
    if request.path.startswith("/api/"):
        return jsonify({"error": "Session expired — reload the page and sign in."}), 401
    return redirect(url_for("login", next=request.path))


@app.route("/login", methods=["GET", "POST"])
def login():
    if not APP_PASSWORD:
        return redirect(url_for("index"))
    error = ""
    if request.method == "POST":
        if hmac.compare_digest(request.form.get("password", ""), APP_PASSWORD):
            session["ok"] = True
            session.permanent = True
            return redirect(request.args.get("next") or url_for("index"))
        error = "That password is not right."
    return render_template_string(LOGIN_PAGE, error=error), (401 if error else 200)


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
    updir = session_dir(sid)
    existing = (get_session(sid) or {}).get("paths") or []
    paths = []


    for fs in request.files.getlist("files"):
        base = os.path.basename(fs.filename or "")
        if not base or base.startswith("~$") or not base.lower().endswith(FILE_EXTENSIONS):
            continue
        dest = os.path.join(updir, f"{len(existing) + len(paths):04d}-{base}")
        save_upload(fs, dest)
        paths.append(dest)

    if not paths:
        return jsonify({"error": "Nothing usable in that selection — looking for "
                                 + ", ".join(FILE_EXTENSIONS) + "."}), 400

    add_to_session(sid, paths, updir)
    all_paths = (get_session(sid) or {}).get("paths") or []
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
                            w.writerow(cols + ["_source_file"])
                            header_written = True
                        continue
                    if all(p(raw) for p in preds):
                        if dedupe:
                            h = hash(tuple(raw))
                            if h in seen:
                                continue
                            seen.add(h)
                        w.writerow([cell(raw, i) for i in take] + [name])
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

LOGIN_PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mahajan &amp; Associates</title>
<style>
  body{margin:0;min-height:100vh;display:grid;place-items:center;background:#f4f3ef;
       color:#16161a;font-family:"Segoe UI",Inter,system-ui,sans-serif;font-size:14px}
  form{background:#fff;border:1px solid #e6e4dc;padding:34px 32px;width:320px}
  h1{margin:0 0 4px;font-size:14px;font-weight:600;letter-spacing:.16em;text-transform:uppercase}
  .sub{font-family:Georgia,serif;font-style:italic;color:#8b8b83;margin-bottom:24px;font-size:13px}
  label{display:block;font-size:9.5px;font-family:ui-monospace,Consolas,monospace;
        text-transform:uppercase;letter-spacing:.1em;color:#8b8b83;margin-bottom:5px}
  input{width:100%;padding:7px 0;border:0;border-bottom:1px solid #d3d0c6;background:transparent;
        font-size:15px;color:#16161a}
  input:focus{outline:0;border-bottom:2px solid #2c4a6e;margin-bottom:-1px}
  button{width:100%;margin-top:22px;padding:10px;border:1px solid #16161a;background:#16161a;
         color:#fff;font-size:12.5px;font-weight:600;letter-spacing:.03em;cursor:pointer}
  button:hover{background:#2c4a6e;border-color:#2c4a6e}
  .err{color:#8f3a26;font-size:12.5px;margin-top:14px}
</style></head>
<body>
  <form method="post">
    <h1>Mahajan &amp; Associates</h1>
    <div class="sub">Registry Search</div>
    <label for="p">Password</label>
    <input id="p" name="password" type="password" autofocus autocomplete="current-password">
    <button type="submit">Sign in</button>
    {% if error %}<div class="err">{{ error }}</div>{% endif %}
  </form>
</body></html>
"""


PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mahajan &amp; Associates — Registry Search</title>
<style>
  :root{
    --ink:#16161a; --ink-soft:#3c3d42; --muted:#8b8b83;
    --paper:#f4f3ef; --sheet:#ffffff; --well:#fbfaf7;
    --line:#e6e4dc; --line-2:#f0eee8; --line-3:#d3d0c6;
    --accent:#2c4a6e; --accent-soft:#eef1f5; --mark:#f7e9ae;
    --warn:#8f3a26; --warn-soft:#faf0ec;
    --mono:ui-monospace,"SF Mono",Menlo,Consolas,"Liberation Mono",monospace;
    --sans:"Segoe UI Variable Text","Segoe UI",Inter,system-ui,-apple-system,sans-serif;
  }
  *{box-sizing:border-box}
  html,body{height:auto}
  body{margin:0;background:var(--paper);color:var(--ink);font-family:var(--sans);
       font-size:14px;line-height:1.5;-webkit-font-smoothing:antialiased;
       font-variant-numeric:tabular-nums}

  /* two horizontal bands */
  .app{display:grid;grid-template-columns:minmax(0,1fr);gap:12px;
       min-height:100vh;padding:0 22px 22px;max-width:1720px;margin:0 auto}
  .sheet{background:var(--sheet);border:1px solid var(--line);display:flex;
         flex-direction:column;min-width:0}
  .results{min-height:420px}

  /* masthead */
  .masthead{display:flex;align-items:baseline;gap:12px;padding:20px 2px 14px;
            border-bottom:1px solid var(--line-3)}
  .masthead h1{margin:0;font-size:16px;font-weight:600;letter-spacing:.16em;
               text-transform:uppercase}
  .masthead h1 .amp{font-family:Georgia,"Times New Roman",serif;font-style:italic;
                    font-weight:400;letter-spacing:0;font-size:17px;
                    color:var(--accent);padding:0 1px}
  .masthead .rule{flex:1;height:1px;background:var(--line-3);align-self:center}
  .signature{font-size:10.5px;color:var(--muted);letter-spacing:.02em;
             white-space:nowrap}
  .signature .by{font-family:var(--mono);font-size:9px;text-transform:uppercase;
                 letter-spacing:.18em}
  .signature .who{font-family:Georgia,"Times New Roman",serif;font-style:italic;
                  font-size:13px;color:var(--ink-soft);letter-spacing:.01em;
                  margin-left:7px}

  /* band 1 */
  .qgrid{display:grid;grid-template-columns:296px minmax(0,1fr)}
  .src{padding:18px 20px;border-right:1px solid var(--line-2)}
  .filters{padding:18px 22px;min-width:0}

  .cap{font-family:var(--mono);font-size:9.5px;text-transform:uppercase;
       letter-spacing:.16em;color:var(--muted);display:flex;align-items:center;
       gap:10px;margin-bottom:12px}
  .cap .n{margin-left:auto;color:var(--accent);letter-spacing:.04em}
  .cap .hint{text-transform:none;letter-spacing:.02em;font-family:var(--sans);
             font-size:11.5px;color:var(--muted)}

  .drop{border:1px dashed var(--line-3);background:var(--well);padding:22px 14px;
        text-align:center;cursor:pointer;display:block;width:100%;font:inherit;
        color:inherit;transition:border-color .12s, background .12s}
  .drop:hover{border-color:var(--accent);background:var(--accent-soft)}
  .drop.over{border-color:var(--accent);background:var(--accent-soft);border-style:solid}
  .drop .big{font-weight:550;font-size:13px;display:block;letter-spacing:.01em}
  .drop .small{font-size:11.5px;color:var(--muted);display:block;margin-top:3px}
  .drop svg{opacity:.4;margin-bottom:8px}

  .chips{margin-top:14px;max-height:112px;overflow:auto;font-family:var(--mono);
         font-size:11px;color:var(--ink-soft)}
  .chip{padding:4px 0;border-bottom:1px solid var(--line-2);display:flex;gap:8px}
  .chip span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .chip em{margin-left:auto;color:var(--muted);font-style:normal;flex:none}

  /* fields */
  label{display:block;font-size:9.5px;font-family:var(--mono);text-transform:uppercase;
        letter-spacing:.1em;color:var(--muted);margin-bottom:5px}
  input[type=text],input[type=number],input[type=date],select{
    width:100%;padding:7px 0;border:0;border-bottom:1px solid var(--line-3);
    background:transparent;font-family:inherit;font-size:13.5px;color:var(--ink);
    border-radius:0;transition:border-color .12s}
  input:hover{border-bottom-color:var(--muted)}
  input:focus,select:focus{outline:0;border-bottom-color:var(--accent);
                           border-bottom-width:2px;margin-bottom:-1px}
  input::placeholder{color:#b6b6ae}
  .drop:focus-visible,button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
  .field{margin-bottom:15px}
  .pair{display:flex;gap:14px}
  .fgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(168px,1fr));gap:0 22px}
  .fgrid .span2{grid-column:span 2}
  .any{margin-bottom:18px;padding-bottom:16px;border-bottom:1px solid var(--line-2)}
  .any input{font-size:15px;padding:6px 0}

  button{font-family:var(--sans);font-size:12.5px;font-weight:600;padding:9px 20px;
         border:1px solid var(--ink);background:var(--ink);color:#fff;cursor:pointer;
         border-radius:0;letter-spacing:.03em;transition:.12s;white-space:nowrap}
  button:hover:not(:disabled){background:var(--accent);border-color:var(--accent)}
  button.ghost{background:transparent;color:var(--ink-soft);border-color:var(--line-3)}
  button.ghost:hover:not(:disabled){background:var(--accent-soft);color:var(--accent);
                                    border-color:var(--accent)}
  button.link{background:none;border:0;color:var(--accent);padding:0;font-size:10.5px;
              font-family:var(--mono);letter-spacing:.06em;text-transform:uppercase;
              border-bottom:1px solid currentColor;font-weight:500}
  button.link:hover:not(:disabled){background:none;color:var(--ink)}
  button:disabled{opacity:.35;cursor:not-allowed}

  .actions{display:flex;gap:10px;align-items:center;padding:13px 20px;
           border-top:1px solid var(--line);background:var(--well)}
  .actions .msg{font-size:11.5px;color:var(--muted);margin-left:auto;
                text-align:right;font-family:var(--mono)}

  /* band 2 */
  .bar{display:flex;align-items:baseline;gap:14px;padding:14px 20px;
       border-bottom:1px solid var(--line);flex-wrap:wrap}
  .bar h2{margin:0;font-family:var(--mono);font-size:9.5px;text-transform:uppercase;
          letter-spacing:.16em;color:var(--muted);font-weight:500}
  .tally{font-size:17px;font-weight:600;letter-spacing:-.015em}
  .tally b{font-weight:600}
  .tally span{color:var(--muted);font-weight:400;font-size:13px}
  .stat{font-family:var(--mono);font-size:10.5px;color:var(--muted);margin-left:auto;
        letter-spacing:.02em}
  .flash{padding:10px 20px;background:var(--warn-soft);color:var(--warn);
         font-size:12.5px;border-bottom:1px solid var(--line)}

  .tablewrap{overflow-x:auto;min-width:0}
  table{border-collapse:separate;border-spacing:0;width:100%;font-size:13px}
  th{position:sticky;top:0;background:var(--sheet);text-align:left;padding:11px 14px;
     border-bottom:1px solid var(--line-3);font-family:var(--mono);font-size:9.5px;
     text-transform:uppercase;letter-spacing:.11em;white-space:nowrap;z-index:2;
     color:var(--muted);font-weight:500}
  td{padding:11px 14px;border-bottom:1px solid var(--line-2);vertical-align:top}
  tbody tr:hover td{background:var(--accent-soft)}
  td.num{text-align:right;font-family:var(--mono);white-space:nowrap;font-size:12.5px}
  td.long{max-width:340px;min-width:190px;color:var(--ink-soft)}
  .clip{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
  .more{font-size:10px;font-family:var(--mono);color:var(--accent);cursor:pointer;
        background:none;border:0;padding:3px 0 0;letter-spacing:.06em;text-transform:uppercase}
  mark{background:var(--mark);color:inherit;padding:0 2px}

  .pager{display:flex;gap:10px;align-items:center;padding:11px 20px;margin-top:auto;
         border-top:1px solid var(--line);font-family:var(--mono);font-size:11px;
         color:var(--muted);letter-spacing:.03em}
  .pager button{padding:5px 12px;font-size:11px}
  .pager select{width:auto;padding:2px 4px;font-size:11px;font-family:var(--mono)}

  .blank{flex:1;min-height:280px;display:grid;place-content:center;text-align:center;padding:44px;
         color:var(--muted);gap:7px;max-width:420px;margin:0 auto}
  .blank .t{font-size:14px;font-weight:600;color:var(--ink-soft);letter-spacing:.02em}
  .blank .s{font-size:12.5px;line-height:1.6}
  .blank .glyph{font-family:var(--mono);font-size:11px;letter-spacing:.3em;
                color:var(--line-3);margin-bottom:4px}
  .err{color:var(--warn)}
  .hide{display:none!important}
  .spin{display:inline-block;width:9px;height:9px;border:1.5px solid var(--line-3);
        border-top-color:var(--accent);border-radius:50%;animation:sp .6s linear infinite;
        vertical-align:0;margin-right:7px}
  @keyframes sp{to{transform:rotate(360deg)}}

  @media (max-width:900px){
    .app{padding:0 14px 14px}
    .qgrid{grid-template-columns:1fr}
    .src{border-right:0;border-bottom:1px solid var(--line-2)}
  }
</style>
</head>
<body>
<div class="app">

  <header class="masthead">
    <h1>Mahajan <span class="amp">&amp;</span> Associates</h1>
    <span class="rule"></span>
    <span class="signature">
      <span class="by">Built by</span><span class="who">Siddhant Mahalank</span>
    </span>
  </header>

  <!-- band 1 — files and filters -->
  <section class="sheet">
    <div class="qgrid">

      <div class="src">
        <div class="cap">Files <span class="n" id="filecount"></span></div>

        <button type="button" class="drop" id="drop">
          <svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor"
               stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round">
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
          <button type="button" class="link hide" id="clearfiles">clear all</button>
        </div>
      </div>

      <div class="filters">
        <div class="cap">Filters
          <span class="hint">every box is optional</span>
          <button type="button" class="link n" id="clearfilters">reset</button>
        </div>

        <div class="any">
          <div class="field" style="margin:0">
            <label for="_all">Any column</label>
            <input type="text" id="_all" placeholder="word, number or name">
          </div>
        </div>

        <div class="fgrid">
          {% for f in filters %}
            {% if f.type == 'text' %}
              <div class="field">
                <label for="f_{{ f.col }}">{{ f.label }}</label>
                <input type="text" id="f_{{ f.col }}" data-col="{{ f.col }}"></div>
            {% elif f.type == 'number' %}
              <div class="field span2"><label>{{ f.label }}</label><div class="pair">
                <input type="number" data-col="{{ f.col }}__min" placeholder="min">
                <input type="number" data-col="{{ f.col }}__max" placeholder="max"></div></div>
            {% elif f.type == 'date' %}
              <div class="field span2"><label>{{ f.label }}</label><div class="pair">
                <input type="date" data-col="{{ f.col }}__from">
                <input type="date" data-col="{{ f.col }}__to"></div></div>
            {% endif %}
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
      <div class="glyph">— — —</div>
      <div class="t" id="blanktitle">Add files to begin</div>
      <div class="s" id="blanksub">Drop the exports above. Nothing is read until you press Search, and leaving every filter empty returns all rows.</div>
    </div>

    <div class="tablewrap hide" id="tablewrap">
      <table><thead id="thead"></thead><tbody id="tbody"></tbody></table>
    </div>

    <div class="pager hide" id="pager">
      <button class="ghost" id="prev" type="button">Prev</button>
      <span id="pageinfo"></span>
      <button class="ghost" id="next" type="button">Next</button>
      <span style="margin-left:auto">rows
        <select id="size"><option>25</option><option selected>50</option>
          <option>100</option><option>250</option></select></span>
    </div>
  </section>
</div>

<script>
const $ = id => document.getElementById(id);
let ROWS = [], COLS = [], LONG = [], PAGE = 1, NFILES = 0, HILITE = "";
const NUMERIC = new Set(["consideration_amt","marketvalue","stampdutypaid","registrationfees",
                         "docno","internaldocumentnumber","srocode","micrno"]);
const PLAIN = new Set(["docno","internaldocumentnumber","srocode","micrno"]);
const OK_EXT = /\.(xls|xlsx|xlsm|csv|html?)$/i;

const esc = s => String(s ?? "").replace(/[&<>"']/g, c =>
  ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

// ---------- sources: one control, files or folders, sent as picked ----------
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

// recurse into dropped folders
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
    $("filecount").textContent = j.files + (j.files === 1 ? " file" : " files");
    $("chips").innerHTML = j.names.map(n => `<div class="chip"><span>${esc(n)}</span></div>`).join("") +
      (j.files > j.names.length ? `<div class="chip"><em>…and ${j.files - j.names.length} more</em></div>` : "");
    $("clearfiles").classList.remove("hide");
    setMsg("not read yet — press Search", false);
    $("searchbtn").disabled = false; $("exportbtn").disabled = false;
    $("blanktitle").textContent = "Ready to search";
    $("blanksub").textContent = "Press Search. Leaving every filter empty returns all rows.";
    $("picker").value = "";
  } catch (e) { setMsg(e.message, true); }
}

$("clearfiles").onclick = async () => {
  await fetch("/api/clear", {method:"POST"});
  NFILES = 0; ROWS = [];
  $("chips").innerHTML = ""; $("filecount").textContent = "";
  $("clearfiles").classList.add("hide");
  $("searchbtn").disabled = true; $("exportbtn").disabled = true;
  $("tablewrap").classList.add("hide"); $("pager").classList.add("hide");
  $("blank").classList.remove("hide"); $("flash").classList.add("hide");
  $("blanktitle").textContent = "Add files to begin";
  $("blanksub").textContent = "Drop the exports above. Nothing is read until you press Search.";
  $("tally").textContent = "—"; $("stat").textContent = "";
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

    ROWS = j.rows; COLS = j.columns; LONG = j.long; HILITE = ($("_all").value || "").trim();
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

const fmt = (col, v) =>
  (NUMERIC.has(col) && !PLAIN.has(col) && v !== "" && !isNaN(v))
    ? Number(v).toLocaleString("en-IN") : v;

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

  $("thead").innerHTML = "<tr>" + COLS.map(c =>
    `<th>${esc(c.replace(/_/g," ").replace("source file","file"))}</th>`).join("") + "</tr>";

  $("tbody").innerHTML = ROWS.slice((PAGE-1)*size, PAGE*size).map(row =>
    "<tr>" + row.map((val, i) => {
      const c = COLS[i], v = mark(esc(fmt(c, val)));
      if (LONG.includes(c))
        return `<td class="long"><div class="clip">${v}</div>` +
          (String(val).length > 90 ? `<button class="more" type="button">more</button>` : "") + `</td>`;
      return `<td class="${NUMERIC.has(c) ? "num" : ""}">${v}</td>`;
    }).join("") + "</tr>").join("");

  document.querySelectorAll(".more").forEach(b => b.onclick = () => {
    b.textContent = b.previousElementSibling.classList.toggle("clip") ? "more" : "less";
  });

  const from = (PAGE-1)*size + 1, to = Math.min(PAGE*size, ROWS.length);
  $("pageinfo").textContent = `${from}–${to} of ${ROWS.length.toLocaleString()}`;
  $("prev").disabled = PAGE <= 1;
  $("next").disabled = to >= ROWS.length;
  $("tablewrap").scrollTop = 0;
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
