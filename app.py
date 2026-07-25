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

PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<titleLawayers — Search</title>
<style>
  :root{
    /* high-contrast, everything legible at a glance */
    --ink:#101418;          /* near-black body text          */
    --ink-2:#2a3138;        /* secondary text, still dark     */
    --muted:#5a6470;        /* labels — dark enough to read   */
    --paper:#eae7df;        /* page                           */
    --sheet:#ffffff;        /* cards                          */
    --well:#f5f3ee;
    --line:#c7c3b8;         /* visible borders                */
    --line-2:#dcd9cf;
    --navy:#173453;         /* structural accent              */
    --navy-2:#22496f;
    --navy-tint:#e5ecf3;
    --gold:#8a5a12;         /* highlight marks, sparing        */
    --gold-tint:#f6e6bf;
    --warn:#8a2f1c;
    --warn-tint:#f6e3dc;
    --sans:"Segoe UI",Roboto,Inter,system-ui,-apple-system,sans-serif;
    --serif:Georgia,"Times New Roman",serif;
    --mono:ui-monospace,"SF Mono",Consolas,"Liberation Mono",monospace;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--paper);color:var(--ink);font-family:var(--sans);
       font-size:16px;line-height:1.55;-webkit-font-smoothing:antialiased;
       font-variant-numeric:tabular-nums}

  .app{display:grid;grid-template-columns:minmax(0,1fr);gap:16px;
       min-height:100vh;padding:0 26px 26px;max-width:1720px;margin:0 auto}
  .sheet{background:var(--sheet);border:1px solid var(--line);
         display:flex;flex-direction:column;min-width:0}

  /* ---------- masthead: both names clearly visible ---------- */
  .masthead{padding:26px 4px 18px;border-bottom:3px double var(--navy)}
  .masthead .firm{font-family:var(--serif);font-size:30px;font-weight:700;
                  color:var(--ink);letter-spacing:.01em;line-height:1.1}
  .masthead .firm .amp{color:var(--navy-2);font-style:italic;padding:0 2px}
  .masthead .strap{display:flex;align-items:baseline;gap:12px;margin-top:8px}
  .masthead .tool{font-family:var(--mono);font-size:12px;text-transform:uppercase;
                  letter-spacing:.22em;color:var(--navy-2);font-weight:600}
  .masthead .by{margin-left:auto;font-size:14px;color:var(--ink-2)}
  .masthead .by b{font-family:var(--serif);font-style:italic;font-weight:700;
                  color:var(--ink);font-size:16px}

  /* ---------- band 1 ---------- */
  .qgrid{display:grid;grid-template-columns:326px minmax(0,1fr)}
  .src{padding:22px 24px;border-right:1px solid var(--line-2)}
  .filters{padding:22px 26px;min-width:0}

  .cap{font-family:var(--mono);font-size:12px;text-transform:uppercase;
       letter-spacing:.14em;color:var(--navy-2);font-weight:700;
       display:flex;align-items:center;gap:10px;margin-bottom:14px}
  .cap .n{margin-left:auto;color:var(--ink);letter-spacing:.02em;
          background:var(--navy-tint);padding:2px 9px;border:1px solid var(--line-2);
          font-size:12px}
  .cap .hint{text-transform:none;letter-spacing:0;font-family:var(--sans);
             font-size:13px;color:var(--muted);font-weight:400}

  .drop{border:2px dashed var(--line);background:var(--well);padding:26px 16px;
        text-align:center;cursor:pointer;display:block;width:100%;font:inherit;
        color:var(--ink);transition:border-color .12s, background .12s}
  .drop:hover{border-color:var(--navy);background:var(--navy-tint)}
  .drop.over{border-color:var(--navy);background:var(--navy-tint);border-style:solid}
  .drop .big{font-weight:600;font-size:16px;display:block}
  .drop .small{font-size:13.5px;color:var(--muted);display:block;margin-top:4px}
  .drop svg{opacity:.55;margin-bottom:9px}

  .chips{margin-top:16px;max-height:150px;overflow:auto;font-size:13.5px;
         color:var(--ink-2)}
  .chip{padding:6px 0;border-bottom:1px solid var(--line-2);display:flex;gap:8px}
  .chip span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .chip em{margin-left:auto;color:var(--muted);font-style:normal;flex:none}

  label{display:block;font-size:12.5px;font-family:var(--mono);text-transform:uppercase;
        letter-spacing:.08em;color:var(--muted);margin-bottom:6px;font-weight:600}
  input[type=text]{
    width:100%;padding:11px 13px;border:1.5px solid var(--line);
    background:#fff;font-family:inherit;font-size:16px;color:var(--ink);border-radius:2px}
  input::placeholder{color:#9aa1a8}
  input:focus,.drop:focus-visible,button:focus-visible{
    outline:2.5px solid var(--navy);outline-offset:1px;border-color:var(--navy)}
  .field{margin-bottom:16px}
  .fgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:0 20px}
  .any{margin-bottom:20px;padding-bottom:18px;border-bottom:2px solid var(--line-2)}
  .any input{font-size:17px;padding:12px 14px}
  .any label{font-size:13px;color:var(--navy-2)}

  button{font-family:var(--sans);font-size:15px;font-weight:600;padding:12px 24px;
         border:1.5px solid var(--navy);background:var(--navy);color:#fff;cursor:pointer;
         border-radius:2px;letter-spacing:.02em;transition:.12s}
  button:hover:not(:disabled){background:var(--navy-2);border-color:var(--navy-2)}
  button.ghost{background:#fff;color:var(--navy);border-color:var(--navy)}
  button.ghost:hover:not(:disabled){background:var(--navy-tint)}
  button.link{background:none;border:0;color:var(--navy-2);padding:0;font-size:13px;
              font-family:var(--sans);text-decoration:underline;text-underline-offset:3px;
              font-weight:600;letter-spacing:0}
  button.link:hover:not(:disabled){background:none;color:var(--ink)}
  button:disabled{opacity:.4;cursor:not-allowed}

  .actions{display:flex;gap:12px;align-items:center;padding:16px 24px;
           border-top:1px solid var(--line);background:var(--well)}
  .actions .msg{font-size:14px;color:var(--ink-2);margin-left:auto;text-align:right}

  /* ---------- band 2: results ---------- */
  .bar{display:flex;align-items:baseline;gap:16px;padding:16px 24px;
       border-bottom:2px solid var(--line);flex-wrap:wrap;background:var(--well)}
  .bar h2{margin:0;font-family:var(--mono);font-size:12px;text-transform:uppercase;
          letter-spacing:.16em;color:var(--navy-2);font-weight:700}
  .tally{font-size:21px;font-weight:700;letter-spacing:-.01em;color:var(--ink)}
  .tally span{color:var(--muted);font-weight:400;font-size:15px}
  .stat{font-family:var(--mono);font-size:13px;color:var(--ink-2);margin-left:auto}
  .flash{padding:12px 24px;background:var(--warn-tint);color:var(--warn);
         font-size:14px;border-bottom:1px solid var(--line);font-weight:500}

  .tablewrap{min-width:0}
  table{border-collapse:separate;border-spacing:0;width:100%;font-size:15px;
        table-layout:fixed}
  /* explicit shares so nothing overflows; party columns kept compact,
     property given the most room, short columns trimmed to what they need */
  col.c-docno{width:7%}
  col.c-registrationdate{width:8%}
  col.c-docname{width:9%}
  col.c-areaname{width:8%}
  col.c-sroname{width:11%}
  col.c-sellerparty{width:15%}
  col.c-purchaserparty{width:15%}
  col.c-propertydescription{width:27%}
  th{position:sticky;top:0;background:var(--navy);text-align:left;padding:13px 16px;
     border-bottom:2px solid var(--navy);font-family:var(--sans);font-size:12.5px;
     text-transform:uppercase;letter-spacing:.06em;white-space:nowrap;z-index:2;
     color:#fff;font-weight:600}
  td{padding:13px 16px;border-bottom:1px solid var(--line-2);vertical-align:top;
     color:var(--ink);line-height:1.5;white-space:normal;
     overflow-wrap:break-word;word-break:break-word}
  tbody tr:nth-child(even) td{background:var(--well)}
  tbody tr:hover td{background:var(--navy-tint)}
  td.party, td.prop{font-size:13.5px;line-height:1.45}
  mark{background:var(--gold-tint);color:var(--ink);padding:0 2px;font-weight:600}

  .pager{display:flex;gap:12px;align-items:center;padding:14px 24px;
         border-top:1px solid var(--line);font-size:14px;color:var(--ink-2)}
  .pager button{padding:8px 15px;font-size:14px}
  .pager select{padding:6px 8px;font-size:14px;border:1.5px solid var(--line);
                background:#fff;color:var(--ink);border-radius:2px}
  .pager b{color:var(--ink)}

  .blank{flex:1;min-height:300px;display:grid;place-content:center;text-align:center;
         padding:48px;color:var(--ink-2);gap:9px;max-width:460px;margin:0 auto}
  .blank .t{font-size:18px;font-weight:700;color:var(--ink)}
  .blank .s{font-size:15px;line-height:1.65;color:var(--muted)}
  .blank .glyph{font-family:var(--serif);font-size:34px;color:var(--line);margin-bottom:2px}
  .err{color:var(--warn);font-weight:600}
  .hide{display:none!important}
  .spin{display:inline-block;width:12px;height:12px;border:2px solid var(--line);
        border-top-color:var(--navy);border-radius:50%;animation:sp .6s linear infinite;
        vertical-align:-1px;margin-right:8px}
  @keyframes sp{to{transform:rotate(360deg)}}

  @media (max-width:920px){
    .app{padding:0 16px 16px}
    .qgrid{grid-template-columns:1fr}
    .src{border-right:0;border-bottom:1px solid var(--line-2)}
  }
</style>
</head>
<body>
<div class="app">

  <header class="masthead">
    <div class="firm">Mahajan <span class="amp">&amp;</span> Associates</div>
    <div class="strap">
      <span class="tool">Registry Search</span>
      <span class="by">Built by <b>Siddhant Mahalank</b></span>
    </div>
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
</div>

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
