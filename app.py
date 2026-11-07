#!/usr/bin/env python3
"""
Registry Search — retired. This app now does one thing: tell visitors the
tool has moved and send them to the new address.

Deploy exactly as before (Render start command `python app.py`, or
`gunicorn app:app`). Only Flask is needed now — lxml, openpyxl, xlrd and
waitress can come out of requirements.txt, though leaving them does no harm.
"""

import os

from flask import Flask, Response, jsonify, render_template_string

# ============================================================================
# CONFIG — edit this block only
# ============================================================================

# Where the tool lives now. Every path on this host shows the notice below
# and then sends the visitor here.
NEW_URL = os.environ.get("NEW_URL", "https://registrysearch.siddhantmahalank.workers.dev/")

# Seconds the notice is shown before the browser moves on. Set to 0 to go
# immediately (the notice then flashes rather than being read).
REDIRECT_DELAY = int(os.environ.get("REDIRECT_DELAY", "5"))

HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "5000"))

# Kept so an existing AdSense verification is not disturbed. Delete the route
# below if the publisher id has moved to the new host.
ADS_TXT = "google.com, pub-5876158560001890, DIRECT, f08c47fec0942fa0\n"

# ============================================================================
# The notice page
# ============================================================================

MOVED_PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>RegistrySearch has moved</title>
<meta name="robots" content="noindex">
<link rel="canonical" href="{{ new_url }}">
<!-- Works even with JavaScript off. -->
<meta http-equiv="refresh" content="{{ delay }};url={{ new_url }}">
<style>
  :root{
    --ink:#111827;--ink-2:#374151;--muted:#6b7280;--paper:#f9fafb;--sheet:#fff;
    --line:#e5e7eb;--navy:#1e3a5f;--navy-2:#2563eb;--navy-3:#1d4ed8;--navy-soft:#93b4d8;
    --sans:"Segoe UI",Roboto,Inter,system-ui,-apple-system,sans-serif;
    --serif:Georgia,"Times New Roman",serif;
    --mono:ui-monospace,"SF Mono",Consolas,monospace;
  }
  *{box-sizing:border-box}
  body{margin:0;min-height:100vh;display:grid;place-items:center;padding:24px;
       background:var(--paper);color:var(--ink);font-family:var(--sans);
       font-size:16px;line-height:1.65;-webkit-font-smoothing:antialiased}
  .card{background:var(--sheet);border:1px solid var(--line);border-radius:12px;
        box-shadow:0 1px 6px rgba(0,0,0,.07);max-width:560px;width:100%;overflow:hidden}
  .top{background:var(--navy);color:#fff;padding:20px 28px;border-bottom:3px solid var(--navy-2)}
  .firm{font-family:var(--serif);font-size:25px;font-weight:700;line-height:1.1}
  .firm .amp{color:var(--navy-soft);font-style:italic;padding:0 2px}
  .body{padding:28px}
  h1{font-family:var(--serif);font-size:26px;color:var(--navy);margin:0 0 10px;line-height:1.25}
  p{margin:0 0 14px;color:var(--ink-2)}
  .newlink{display:block;font-family:var(--mono);font-size:14px;word-break:break-all;
           background:#f3f4f6;border:1px solid var(--line);border-radius:7px;
           padding:12px 14px;margin:0 0 18px;color:var(--navy-2);
           text-decoration:none}
  .newlink:hover{background:#eff6ff;border-color:var(--navy-2)}
  .cta{display:inline-block;background:var(--navy-2);color:#fff;text-decoration:none;
       font-weight:700;padding:11px 24px;border-radius:7px}
  .cta:hover{background:var(--navy-3)}
  .count{font-size:14px;color:var(--muted);margin:16px 0 0}
  .count b{color:var(--ink);font-variant-numeric:tabular-nums}
  .mr{font-size:14.5px;color:var(--ink-2);background:#f3f4f6;
      border-left:3px solid var(--navy-soft);padding:10px 14px;margin:18px 0 0;
      border-radius:0 4px 4px 0}
</style>
</head>
<body>
  <main class="card">
    <div class="top"><span class="firm">Registry<span class="amp">Search</span></span></div>
    <div class="body">
      <h1>We have shifted</h1>
      <p>RegistrySearch has moved to a new address. This page will take you
         there automatically.</p>
      <a class="newlink" href="{{ new_url }}">{{ new_url }}</a>
      <a class="cta" href="{{ new_url }}">Go there now →</a>
      <p class="count" id="count">Redirecting in <b id="n">{{ delay }}</b> seconds…</p>
      <div class="mr">RegistrySearch आता नवीन पत्त्यावर उपलब्ध आहे. काही क्षणात
        तुम्हाला तिकडे पाठवले जाईल.</div>
    </div>
  </main>
<script>
  var url = {{ new_url | tojson }};
  var left = {{ delay }};
  var n = document.getElementById("n");
  if (left > 0) {
    var t = setInterval(function () {
      left -= 1;
      if (n) n.textContent = left;
      if (left <= 0) { clearInterval(t); location.replace(url); }
    }, 1000);
  } else {
    location.replace(url);
  }
</script>
</body>
</html>
"""

# ============================================================================
# App — every path leads to the notice
# ============================================================================

app = Flask(__name__)


def moved():
    html = render_template_string(MOVED_PAGE, new_url=NEW_URL, delay=REDIRECT_DELAY)
    # mimetype, not content_type: Flask appends "; charset=utf-8" itself, and
    # spelling it out here produces a duplicated charset parameter.
    resp = Response(html, mimetype="text/html")
    resp.headers["X-Robots-Tag"] = "noindex"
    resp.headers["Cache-Control"] = "no-store"
    # Tell crawlers and link checkers where this went, without redirecting
    # the human before they have read the notice.
    resp.headers["Link"] = f'<{NEW_URL}>; rel="canonical"'
    return resp


@app.route("/healthz")
def healthz():
    """Left in place so an existing uptime pinger keeps getting a 200. Nothing
    here needs keeping warm any more — the pinger can be switched off."""
    return "ok", 200


@app.route("/ads.txt")
def ads_txt():
    return Response(ADS_TXT, mimetype="text/plain")


@app.route("/api/<path:_rest>", methods=["GET", "POST", "PUT", "DELETE"])
def api_gone(_rest):
    """410 rather than the notice page: a stale cached copy of the old
    JavaScript must not be able to upload files to a retired service."""
    return jsonify({
        "error": "RegistrySearch has moved.",
        "url": NEW_URL,
    }), 410


@app.route("/", defaults={"_path": ""})
@app.route("/<path:_path>")
def catch_all(_path):
    """Any old link — /guide, /about, /privacy, /contact, a bookmarked deep
    link — lands on the notice rather than a 404."""
    return moved()


if __name__ == "__main__":
    print(f"\n  RegistrySearch redirect service — forwarding to {NEW_URL}")
    print(f"  Listening on http://{HOST}:{PORT}\n")
    try:
        from waitress import serve
        serve(app, host=HOST, port=PORT, threads=2)
    except ImportError:
        app.run(host=HOST, port=PORT, threaded=True, debug=False)
