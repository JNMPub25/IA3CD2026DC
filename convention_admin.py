#!/usr/bin/env python3
"""
Convention Admin Dashboard — Browser-based setup tool
=====================================================
Replaces the terminal-based district3_setup.py with a visual interface.
Same data files, same QR/display generation — just open in your browser.

USAGE:
    python convention_admin.py
    → Opens http://localhost:8026 in your default browser

REQUIREMENTS (already installed):
    pip install qrcode[pil]
"""

import os, json, csv, sys, re, io, base64, urllib.parse, webbrowser, threading
from datetime import datetime, timedelta
from pathlib import Path
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# ── QR code library (graceful fallback) ──────────────────────────────────
try:
    import qrcode
    QRCODE_AVAILABLE = True
except ImportError:
    QRCODE_AVAILABLE = False

# ── Folder layout (same as district3_setup.py) ──────────────────────────
BASE_DIR    = Path(__file__).parent
DATA_DIR    = BASE_DIR / "setup_data"
QR_DIR      = BASE_DIR / "qr_codes"
DISPLAY_DIR = BASE_DIR / "projection_displays"

for d in [DATA_DIR, QR_DIR, DISPLAY_DIR]:
    d.mkdir(exist_ok=True)

NOMINEES_FILE    = DATA_DIR / "nominees.json"
DELEGATES_FILE   = DATA_DIR / "delegates.json"
SURRENDERED_FILE = DATA_DIR / "surrendered_delegates.json"

# ── Election definitions ─────────────────────────────────────────────────
# Defaults — overridden by setup_data/elections.json if it exists
DEFAULT_ELECTIONS = {
    "scc-w":   {"key": "scc-w",   "label": "SCC — Women",              "seats": 4,  "type": "ranked"},
    "scc-m":   {"key": "scc-m",   "label": "SCC — Men / Non-Binary",   "seats": 4,  "type": "ranked"},
    "dei":     {"key": "dei",     "label": "DEI Committee Chair",       "seats": 1,  "type": "ranked"},
    "scc-com": {"key": "scc-com", "label": "State Convention Committee","seats": 14, "type": "slate"},
}
ELECTIONS_FILE = DATA_DIR / "elections.json"

def get_elections():
    """Load elections from JSON file, falling back to defaults."""
    return load_json(ELECTIONS_FILE, DEFAULT_ELECTIONS)

def save_elections(elections):
    save_json(ELECTIONS_FILE, elections)

# ELECTIONS initialized after load_json is defined (see below)

BALLOT_BASE_URL   = "https://jnmpub25.github.io/IA3CD2026DC/ballot.html"
PRACTICE_BASE_URL = "https://jnmpub25.github.io/IA3CD2026DC/practice_ballot.html"

PRACTICE_CANDIDATES = [
    {"letter": "A", "name": "Adams, Carol"},
    {"letter": "B", "name": "Bennett, Diane"},
    {"letter": "C", "name": "Chen, Megan"},
    {"letter": "D", "name": "Davis, Rachel"},
    {"letter": "E", "name": "Evans, Patricia"},
    {"letter": "F", "name": "Flores, Sandra"},
]

PORT = 8026

# ── Apps Script URLs (for live voting controls + setup push) ─────────────
# Production ballot Apps Script
APPS_SCRIPT_PROD_URL = "https://script.google.com/macros/s/AKfycbw8jmqZN14Q0I-Q7pDaQwWxS4pkJHTmCXNvFAuGi9MEXcNn5doiW8-CZ-PXGkQ4lK_u/exec"
# Test ballot Apps Script
APPS_SCRIPT_TEST_URL = "https://script.google.com/macros/s/AKfycbyteBS-_d70RzznyC5q85z8QGZbMvCCkNm-vdvrOqfLuA43zAxUxqR5FghPnZwjbAfe-Q/exec"
# Production credentials Apps Script
CREDENTIALS_PROD_URL = "https://script.google.com/macros/s/AKfycbwJPg9pyoB6h-dikU2I21_Hc0H9MBTZwdIfXIkTRACFPVv_zaSeOkiWgQI9-kGc2c0C/exec"
# Test credentials Apps Script
CREDENTIALS_TEST_URL = "https://script.google.com/macros/s/AKfycbwurkU9Z9wmNUaaKKFAuHTLUvc4rmFRNkkCP29cuSBhH4s80_fO2_MkxQRp0iUYcrms/exec"

# Default to PROD for convention day; switch to TEST for testing
APPS_SCRIPT_URL  = APPS_SCRIPT_PROD_URL
CREDENTIALS_URL  = CREDENTIALS_PROD_URL

# ═════════════════════════════════════════════════════════════════════════
#  DATA HELPERS
# ═════════════════════════════════════════════════════════════════════════

def load_json(path, default):
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return default

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

# Now that load_json is defined, initialize ELECTIONS
ELECTIONS = get_elections()

def get_nominees():
    return load_json(NOMINEES_FILE, {e["key"]: [] for e in get_elections().values()})

def get_delegates():
    return load_json(DELEGATES_FILE, [])

def get_surrendered():
    return load_json(SURRENDERED_FILE, [])


# ═════════════════════════════════════════════════════════════════════════
#  APPS SCRIPT PROXY (avoids CORS issues from browser)
# ═════════════════════════════════════════════════════════════════════════

def apps_script_post(payload, use_test=False):
    """POST JSON to the ballot Apps Script and return the response."""
    import urllib.request
    url = APPS_SCRIPT_TEST_URL if use_test else APPS_SCRIPT_URL
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    try:
        # Apps Script redirects (302) on web app calls — follow redirects
        resp = urllib.request.urlopen(req, timeout=15)
        return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        # Apps Script returns 302 redirect for web apps — follow it
        if e.code == 302:
            redirect_url = e.headers.get("Location", "")
            if redirect_url:
                req2 = urllib.request.Request(redirect_url)
                resp2 = urllib.request.urlopen(req2, timeout=15)
                return json.loads(resp2.read().decode("utf-8"))
        return {"status": "error", "message": f"HTTP {e.code}: {str(e)}"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

def apps_script_get(election_key="", use_test=False):
    """GET status from the ballot Apps Script."""
    import urllib.request
    url = APPS_SCRIPT_TEST_URL if use_test else APPS_SCRIPT_URL
    if election_key:
        url += f"?election={election_key}"
    try:
        resp = urllib.request.urlopen(url, timeout=15)
        return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return {"status": "error", "message": str(e)}


# ═════════════════════════════════════════════════════════════════════════
#  QR CODE GENERATION
# ═════════════════════════════════════════════════════════════════════════

def generate_qr_base64(url):
    """Generate a QR code PNG and return as base64 data URI."""
    if not QRCODE_AVAILABLE:
        return None
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10, border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()

def save_qr_png(url, filename):
    """Save QR code to qr_codes/ directory."""
    if not QRCODE_AVAILABLE:
        return None
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10, border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    path = QR_DIR / filename
    img.save(str(path))
    return str(path)

# ═════════════════════════════════════════════════════════════════════════
#  PROJECTION DISPLAY GENERATION
# ═════════════════════════════════════════════════════════════════════════

def write_projection_html(path, election, nominees, url, phase2_minutes=None,
                          timer_minutes=None):
    candidate_rows = "".join(
        f'<li><span class="num">{n["letter"]}</span>{n["name"]}</li>'
        for n in nominees
    )
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{election['label']} — Ballot Display</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: #003366; color: #ffffff;
    font-family: 'Segoe UI', Arial, sans-serif;
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; min-height: 100vh; padding: 2rem;
  }}
  .header {{ text-align: center; margin-bottom: 2rem; }}
  .header h1 {{ font-size: 3rem; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; }}
  .header p {{ font-size: 1.4rem; opacity: 0.85; margin-top: 0.4rem; }}
  .content {{ display: flex; gap: 4rem; align-items: flex-start; justify-content: center; flex-wrap: wrap; width: 100%; max-width: 1200px; }}
  .candidates {{ flex: 1; min-width: 300px; }}
  .candidates h2 {{ font-size: 1.4rem; text-transform: uppercase; letter-spacing: 0.1em; opacity: 0.7; margin-bottom: 1rem; border-bottom: 2px solid rgba(255,255,255,0.3); padding-bottom: 0.4rem; }}
  .candidates ol {{ list-style: none; padding: 0; }}
  .candidates li {{ font-size: 1.8rem; padding: 0.5rem 0; border-bottom: 1px solid rgba(255,255,255,0.15); display: flex; align-items: center; gap: 1rem; }}
  .num {{ display: inline-flex; align-items: center; justify-content: center; background: rgba(255,255,255,0.2); border-radius: 50%; width: 2.2rem; height: 2.2rem; font-size: 1.1rem; font-weight: 700; flex-shrink: 0; }}
  .qr-panel {{ display: flex; flex-direction: column; align-items: center; gap: 1rem; }}
  .qr-panel h2 {{ font-size: 1.4rem; text-transform: uppercase; letter-spacing: 0.1em; opacity: 0.7; }}
  #qrcode canvas, #qrcode img {{ border: 8px solid white; border-radius: 8px; }}
  .url-box {{ background: rgba(255,255,255,0.15); border-radius: 6px; padding: 0.8rem 1.2rem; font-size: 0.85rem; word-break: break-all; max-width: 340px; text-align: center; opacity: 0.8; }}
  .seats-badge {{ display: inline-block; background: #FFD700; color: #003366; font-weight: 700; font-size: 1rem; padding: 0.3rem 1rem; border-radius: 20px; margin-top: 0.5rem; }}
  .phase2-badge {{ display: inline-block; background: #c0392b; color: #fff; font-weight: 700; font-size: 1.1rem; padding: 0.35rem 1.2rem; border-radius: 20px; margin-top: 0.6rem; letter-spacing: 0.06em; }}
</style>
</head>
<body>
<div class="header">
  <h1>3rd Congressional District</h1>
  <p>Iowa Democratic Party — District Convention {datetime.now().year}</p>
  <div class="seats-badge">{election['label'].upper()} &nbsp;|&nbsp; {election['seats']} Seat{'s' if election['seats'] != 1 else ''}</div>
  {f'<div class="phase2-badge">⏱ PHASE 2 — {phase2_minutes} MINUTES</div>' if phase2_minutes else ''}
  {f'<div class="phase2-badge" style="background:#1a7a4a;">⏱ TIMER — {timer_minutes} MIN WINDOW</div>' if timer_minutes else ''}
</div>
<div class="content">
  <div class="candidates">
    <h2>Candidates on Ballot</h2>
    <ol>{candidate_rows}</ol>
  </div>
  <div class="qr-panel">
    <h2>Scan to Vote</h2>
    <div id="qrcode"></div>
    <div class="url-box">{url}</div>
  </div>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
<script>
  new QRCode(document.getElementById("qrcode"), {{
    text: "{url}",
    width: 300, height: 300,
    colorDark: "#000000", colorLight: "#ffffff",
    correctLevel: QRCode.CorrectLevel.M
  }});
</script>
</body>
</html>"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return str(path)


def write_practice_display_html(path, url):
    candidate_rows = "".join(
        f'<li><span class="num">{c["letter"]}</span>{c["name"]}</li>'
        for c in PRACTICE_CANDIDATES
    )
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Delegate Practice Ballot — Display</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: #3e2700; color: #ffffff; font-family: 'Segoe UI', Arial, sans-serif;
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    min-height: 100vh; padding: 2rem; }}
  .header {{ text-align: center; margin-bottom: 2rem; }}
  .header .practice-label {{ display: inline-block; background: #f9a825; color: #3e2700;
    font-size: 1rem; font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase;
    padding: 0.3rem 1rem; border-radius: 4px; margin-bottom: 0.75rem; }}
  .header h1 {{ font-size: 2.8rem; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; }}
  .header p {{ font-size: 1.3rem; opacity: 0.85; margin-top: 0.4rem; }}
  .content {{ display: flex; gap: 4rem; align-items: flex-start; justify-content: center; flex-wrap: wrap; width: 100%; }}
  .candidates {{ flex: 1; min-width: 320px; max-width: 520px; }}
  .candidates h2 {{ font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.1em; opacity: 0.7; margin-bottom: 0.75rem; color: #f9a825; }}
  .candidates ul {{ list-style: none; }}
  .candidates li {{ display: flex; align-items: center; gap: 1rem; font-size: 1.8rem; padding: 0.45rem 0;
    border-bottom: 1px solid rgba(249,168,37,0.2); font-style: italic; opacity: 0.9; }}
  .candidates li:last-child {{ border-bottom: none; }}
  .num {{ display: inline-flex; align-items: center; justify-content: center; width: 2.4rem; height: 2.4rem;
    border-radius: 50%; background: #f9a825; color: #3e2700; font-size: 1.1rem; font-weight: 700; font-style: normal; flex-shrink: 0; }}
  .qr-panel {{ display: flex; flex-direction: column; align-items: center; gap: 1rem; min-width: 260px; }}
  .qr-panel h2 {{ font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.1em; opacity: 0.7; color: #f9a825; }}
  #qrcode canvas, #qrcode img {{ border-radius: 8px; }}
  .url-text {{ font-size: 0.78rem; opacity: 0.55; word-break: break-all; max-width: 280px; text-align: center; }}
  .footer-note {{ margin-top: 2.5rem; font-size: 1rem; opacity: 0.6; text-align: center; letter-spacing: 0.03em; }}
</style>
</head>
<body>
  <div class="header">
    <div class="practice-label">Practice Session — Fictional Candidates Only</div>
    <h1>Ranked-Choice Practice Ballot</h1>
    <p>3rd Congressional District Convention — Before We Begin</p>
  </div>
  <div class="content">
    <div class="candidates">
      <h2>Practice Candidates</h2>
      <ul>{candidate_rows}</ul>
    </div>
    <div class="qr-panel">
      <h2>Scan to Practice</h2>
      <div id="qrcode"></div>
      <p class="url-text">{url}</p>
    </div>
  </div>
  <p class="footer-note">Enter any number as your delegate number &middot; Rank the candidates &middot; No votes are recorded</p>
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
<script>
  new QRCode(document.getElementById("qrcode"), {{
    text: "{url}", width: 240, height: 240,
    colorDark: "#000000", colorLight: "#ffffff",
    correctLevel: QRCode.CorrectLevel.M
  }});
</script>
</body>
</html>"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return str(path)


# ═════════════════════════════════════════════════════════════════════════
#  URL BUILDING
# ═════════════════════════════════════════════════════════════════════════

def build_ballot_url(election_key, nominees, duration=None, admin=False):
    """Build a ballot URL with candidate parameters."""
    election = get_elections()[election_key]
    candidate_param = "-".join(
        f"{n['letter']}-{urllib.parse.quote(n['name'], safe='')}"
        for n in nominees
    )
    url = (f"{BALLOT_BASE_URL}"
           f"?election={election_key}"
           f"&candidates={candidate_param}"
           f"&seats={election['seats']}")
    if duration:
        url += f"&duration={duration}"
    if admin:
        url += "&admin=1"
    return url

def build_assisted_url(election_key, nominees, duration=None):
    """Build an assisted-mode ballot URL."""
    url = build_ballot_url(election_key, nominees, duration)
    url += "&mode=assisted"
    return url


# ═════════════════════════════════════════════════════════════════════════
#  HTML DASHBOARD
# ═════════════════════════════════════════════════════════════════════════


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Convention Admin Dashboard</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--navy:#003366;--navy-light:#004a8f;--gold:#FFD700;--bg:#f0f2f5;--card:#fff;--text:#1a1a2e;--text-muted:#666;--border:#ddd;--success:#27ae60;--danger:#c0392b;--warning:#f39c12}
html,body{height:100%;font-family:'Segoe UI',Arial,sans-serif;font-size:15px;background:var(--bg);color:var(--text)}
.app{display:flex;height:100vh}
.sidebar{width:240px;background:var(--navy);color:#fff;padding:1.5rem 0;display:flex;flex-direction:column;flex-shrink:0}
.sidebar h1{font-size:1rem;text-align:center;padding:0 1rem .25rem;letter-spacing:.04em}
.sidebar .subtitle{font-size:.75rem;text-align:center;opacity:.6;padding-bottom:1.25rem;border-bottom:1px solid rgba(255,255,255,.15);margin-bottom:.5rem}
.nav-item{display:flex;align-items:center;gap:.75rem;padding:.7rem 1.25rem;cursor:pointer;transition:background .15s;font-size:.9rem;border:none;background:none;color:#fff;width:100%;text-align:left}
.nav-item:hover{background:rgba(255,255,255,.1)}
.nav-item.active{background:var(--navy-light);border-left:3px solid var(--gold)}
.nav-item .icon{font-size:1.1rem;width:1.5rem;text-align:center}
.main{flex:1;overflow-y:auto;padding:2rem}
.main h2{font-size:1.5rem;color:var(--navy);margin-bottom:1.25rem}
.card{background:var(--card);border-radius:10px;padding:1.5rem;box-shadow:0 1px 6px rgba(0,0,0,.08);margin-bottom:1.25rem}
.card h3{font-size:1.1rem;color:var(--navy);margin-bottom:.75rem}
.election-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:1rem;margin-bottom:1.5rem}
.election-card{background:var(--card);border-radius:10px;padding:1.25rem;box-shadow:0 1px 6px rgba(0,0,0,.08);border-left:4px solid var(--navy)}
.election-card h4{font-size:1rem;color:var(--navy);margin-bottom:.5rem}
.election-card .meta{font-size:.85rem;color:var(--text-muted)}
.election-card .count{font-size:1.8rem;font-weight:700;color:var(--navy)}
table{width:100%;border-collapse:collapse;font-size:.9rem}
th{background:var(--navy);color:#fff;padding:.6rem .75rem;text-align:left;font-weight:600}
td{padding:.55rem .75rem;border-bottom:1px solid var(--border)}
tr:hover td{background:#f8f9fa}
.form-row{display:flex;gap:.75rem;align-items:flex-end;margin-bottom:.75rem;flex-wrap:wrap}
.form-group{display:flex;flex-direction:column;gap:.25rem}
.form-group label{font-size:.8rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em}
input,select{padding:.5rem .75rem;border:1px solid var(--border);border-radius:6px;font-size:.9rem;font-family:inherit}
input:focus,select:focus{outline:none;border-color:var(--navy);box-shadow:0 0 0 2px rgba(0,51,102,.15)}
.btn{display:inline-flex;align-items:center;gap:.4rem;padding:.5rem 1rem;border:none;border-radius:6px;cursor:pointer;font-size:.85rem;font-weight:600;font-family:inherit;transition:opacity .15s}
.btn:hover{opacity:.85}
.btn-primary{background:var(--navy);color:#fff}
.btn-success{background:var(--success);color:#fff}
.btn-danger{background:var(--danger);color:#fff}
.btn-warning{background:var(--warning);color:#fff}
.btn-sm{padding:.3rem .65rem;font-size:.8rem}
.btn-outline{background:transparent;border:1px solid var(--border);color:var(--text)}
.badge{display:inline-block;padding:.15rem .5rem;border-radius:10px;font-size:.75rem;font-weight:600}
.badge-info{background:#e3f2fd;color:#1565c0}
.badge-success{background:#e8f5e9;color:#2e7d32}
.badge-warning{background:#fff8e1;color:#f57f17}
.tab-bar{display:flex;gap:0;border-bottom:2px solid var(--border);margin-bottom:1.25rem}
.tab-btn{padding:.6rem 1.25rem;border:none;background:none;cursor:pointer;font-size:.9rem;font-weight:600;color:var(--text-muted);border-bottom:2px solid transparent;margin-bottom:-2px;font-family:inherit}
.tab-btn.active{color:var(--navy);border-bottom-color:var(--navy)}
.toast-container{position:fixed;top:1rem;right:1rem;z-index:1000;display:flex;flex-direction:column;gap:.5rem}
.toast{padding:.75rem 1.25rem;border-radius:8px;color:#fff;font-size:.9rem;box-shadow:0 4px 12px rgba(0,0,0,.2);animation:slideIn .3s ease;max-width:400px}
.toast-success{background:var(--success)}.toast-error{background:var(--danger)}.toast-info{background:var(--navy)}
@keyframes slideIn{from{transform:translateX(100%);opacity:0}to{transform:translateX(0);opacity:1}}
.section{display:none}.section.active{display:block}
.qr-preview{display:flex;gap:2rem;align-items:flex-start;flex-wrap:wrap;margin-top:1rem}
.qr-preview img{border:2px solid var(--border);border-radius:8px;max-width:200px}
.url-display{background:#f8f9fa;border:1px solid var(--border);border-radius:6px;padding:.75rem 1rem;word-break:break-all;font-size:.82rem;font-family:'Consolas','Courier New',monospace;margin:.5rem 0;max-width:500px;cursor:pointer}
.status-bar{display:flex;gap:1.5rem;padding:.75rem 1.25rem;background:#e8f0fe;border-radius:8px;margin-bottom:1.25rem;font-size:.85rem;flex-wrap:wrap}
.status-item{display:flex;align-items:center;gap:.4rem}
.status-dot{width:8px;height:8px;border-radius:50%}
.status-dot.green{background:var(--success)}.status-dot.red{background:var(--danger)}.status-dot.yellow{background:var(--warning)}
.drop-zone{border:2px dashed var(--border);border-radius:8px;padding:2rem;text-align:center;color:var(--text-muted);cursor:pointer;transition:all .2s;margin-bottom:1rem}
.drop-zone:hover,.drop-zone.drag-over{border-color:var(--navy);background:#f0f4ff;color:var(--navy)}
.drop-zone input[type="file"]{display:none}
.results-panel{background:#f0faf0;border:1px solid #c8e6c9;border-radius:8px;padding:1.25rem;margin-top:1rem}
.results-panel h4{color:var(--success);margin-bottom:.75rem}
.file-link{display:flex;align-items:center;gap:.5rem;padding:.3rem 0;font-size:.85rem}
.nominee-list{display:flex;flex-wrap:wrap;gap:.5rem;margin:.75rem 0}
.nominee-pill{display:inline-flex;align-items:center;gap:.5rem;background:#e3f2fd;border:1px solid #bbdefb;border-radius:20px;padding:.35rem .5rem .35rem .75rem;font-size:.85rem}
.nominee-pill .letter{font-weight:700;color:var(--navy)}
.nominee-pill .remove-btn{background:none;border:none;cursor:pointer;color:var(--danger);font-size:1rem;line-height:1;padding:0 .2rem;opacity:.6}
.nominee-pill .remove-btn:hover{opacity:1}
</style>
</head>
<body>
<div class="app">
  <div class="sidebar">
    <h1>3rd CD Convention</h1>
    <div class="subtitle">Admin Dashboard</div>
    <button class="nav-item active" onclick="showSection('dashboard')" data-section="dashboard"><span class="icon">&#9776;</span> Dashboard</button>
    <button class="nav-item" onclick="showSection('elections')" data-section="elections"><span class="icon">&#9881;</span> Elections</button>
    <button class="nav-item" onclick="showSection('nominees')" data-section="nominees"><span class="icon">&#9733;</span> Nominees</button>
    <button class="nav-item" onclick="showSection('delegates')" data-section="delegates"><span class="icon">&#128101;</span> Delegates</button>
    <button class="nav-item" onclick="showSection('surrendered')" data-section="surrendered"><span class="icon">&#9888;</span> Surrendered</button>
    <button class="nav-item" onclick="showSection('generate')" data-section="generate"><span class="icon">&#9638;</span> QR / Displays</button>
    <button class="nav-item" onclick="showSection('livecontrol')" data-section="livecontrol"><span class="icon">&#9654;</span> Live Controls</button>
    <button class="nav-item" onclick="showSection('tabulate')" data-section="tabulate"><span class="icon">&#9879;</span> Tabulation</button>
    <button class="nav-item" onclick="showSection('urls')" data-section="urls"><span class="icon">&#128279;</span> All URLs</button>
  </div>
  <div class="main">
    <div class="toast-container" id="toast-container"></div>

    <!-- DASHBOARD -->
    <div class="section active" id="sec-dashboard">
      <h2>Convention Dashboard</h2>
      <div class="status-bar" id="status-bar"></div>
      <div class="election-grid" id="dashboard-grid"></div>
      <div class="card"><h3>Quick Actions</h3>
        <div style="display:flex;gap:.75rem;flex-wrap:wrap;">
          <button class="btn btn-primary" onclick="showSection('nominees')">Manage Nominees</button>
          <button class="btn btn-success" onclick="showSection('generate')">Generate QR Codes</button>
          <button class="btn btn-warning" onclick="showSection('surrendered')">Record Surrender</button>
          <button class="btn btn-outline" onclick="showSection('livecontrol')">Live Controls</button>
        </div>
      </div>
    </div>

    <!-- ELECTIONS -->
    <div class="section" id="sec-elections">
      <h2>Election Setup</h2>
      <p style="color:var(--text-muted);margin-bottom:1rem;">Define elections for the convention. Changes update nominee lists, QR generators, and live controls.</p>
      <div class="card"><h3>Add / Edit Election</h3>
        <div class="form-row">
          <div class="form-group"><label>Key (URL id)</label><input type="text" id="el-key" placeholder="e.g. scc-w" style="width:140px;"></div>
          <div class="form-group"><label>Election Name</label><input type="text" id="el-label" placeholder="e.g. SCC — Women" style="width:260px;"></div>
          <div class="form-group"><label>Seats</label><input type="number" id="el-seats" value="1" min="1" style="width:70px;"></div>
          <div class="form-group"><label>Type</label>
            <select id="el-type" style="width:160px;">
              <option value="ranked">Ranked Choice (RCV)</option>
              <option value="slate">Slate Vote (Yes/No)</option>
              <option value="runoff">Runoff / Tiebreaker</option>
            </select>
          </div>
          <button class="btn btn-primary" onclick="saveElection()">Save</button>
        </div>
        <p style="font-size:.78rem;color:var(--text-muted);margin-top:.5rem;">Key is used in ballot URLs and Sheet tab names. Once set, avoid changing it.</p>
      </div>
      <div class="card"><h3>Current Elections</h3>
        <table><thead><tr><th>Key</th><th>Name</th><th>Seats</th><th>Type</th><th>Nominees</th><th></th></tr></thead>
        <tbody id="elections-table"></tbody></table>
      </div>
    </div>

    <!-- NOMINEES -->
    <div class="section" id="sec-nominees">
      <h2>Nominee Management</h2>
      <div class="tab-bar" id="nominee-tabs"></div>
      <div id="nominee-content"></div>
    </div>

    <!-- DELEGATES -->
    <div class="section" id="sec-delegates">
      <h2>Delegate Roster</h2>
      <div class="card"><h3>Import from CSV</h3>
        <div class="drop-zone" id="csv-drop" onclick="document.getElementById('csv-file').click()">
          <input type="file" id="csv-file" accept=".csv" onchange="handleCSVSelect(this)">
          <p><strong>Click to select</strong> or drag-and-drop a CSV file</p>
          <p style="font-size:.8rem;margin-top:.4rem;">You'll map columns after the file loads</p>
        </div>
        <div id="csv-mapper" style="display:none;"></div>
      </div>
      <div class="card"><h3>Add Single Delegate</h3>
        <div class="form-row">
          <div class="form-group"><label>Delegate ID</label><input type="text" id="add-del-id" placeholder="D-6001" style="width:140px;"></div>
          <div class="form-group"><label>Last Name</label><input type="text" id="add-del-last" style="width:160px;"></div>
          <div class="form-group"><label>First Name</label><input type="text" id="add-del-first" style="width:160px;"></div>
          <div class="form-group"><label>Email</label><input type="text" id="add-del-email" style="width:200px;"></div>
          <div class="form-group"><label>Phone</label><input type="text" id="add-del-phone" style="width:140px;"></div>
          <button class="btn btn-primary" onclick="addDelegate()">Add</button>
        </div>
      </div>
      <div class="card">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:.75rem;">
          <h3 id="delegate-count">Roster (0 loaded)</h3>
          <input type="text" id="delegate-search" placeholder="Search..." style="width:200px;" oninput="filterDelegates()">
        </div>
        <div style="max-height:400px;overflow-y:auto;">
          <table><thead><tr><th>ID</th><th>Name</th><th>Email</th><th>Phone</th></tr></thead>
          <tbody id="delegate-table"></tbody></table>
        </div>
      </div>
    </div>

    <!-- SURRENDERED -->
    <div class="section" id="sec-surrendered">
      <h2>Surrendered Credentials</h2>
      <div class="card"><h3>Record Surrender</h3>
        <div class="form-row">
          <div class="form-group"><label>Delegate ID</label><input type="text" id="surr-id" placeholder="D-6001" style="width:160px;"></div>
          <div class="form-group"><label>Status</label>
            <select id="surr-status" style="width:160px;">
              <option value="surrendered">Surrendered</option>
              <option value="non-issued">Non-Issued</option>
            </select>
          </div>
          <div class="form-group"><label>Time (blank=now)</label><input type="text" id="surr-time" placeholder="YYYY-MM-DD HH:MM:SS" style="width:200px;"></div>
          <div class="form-group"><label>Note</label><input type="text" id="surr-note" style="width:200px;"></div>
          <button class="btn btn-warning" onclick="recordSurrender()">Record</button>
        </div>
      </div>
      <div class="card"><h3 id="surr-count">Surrendered List (0)</h3>
        <table><thead><tr><th>Delegate</th><th>Name</th><th>Status</th><th>Time</th><th>Note</th><th></th></tr></thead>
        <tbody id="surr-table"></tbody></table>
      </div>
    </div>

    <!-- QR / DISPLAYS -->
    <div class="section" id="sec-generate">
      <h2>Generate QR Codes &amp; Projection Displays</h2>
      <div class="card"><h3>Election Ballot QR + Display</h3>
        <div class="form-row">
          <div class="form-group"><label>Election</label>
            <select id="gen-election" style="width:260px;" onchange="updateGenPreview()"><option value="">-- Select --</option></select>
          </div>
          <div class="form-group"><label>Include Timer?</label>
            <select id="gen-timer" style="width:140px;" onchange="toggleTimerOptions()"><option value="no">No</option><option value="yes">Yes</option></select>
          </div>
          <div class="form-group" id="gen-duration-group" style="display:none;">
            <label>Duration (min)</label><input type="number" id="gen-duration" style="width:100px;" min="1">
          </div>
          <button class="btn btn-success" onclick="generateQR()">Generate</button>
        </div>
        <div id="gen-candidate-preview" style="margin-top:.5rem;font-size:.85rem;color:var(--text-muted);"></div>
        <div id="gen-results" style="display:none;"></div>
      </div>
      <div class="card"><h3>Practice Ballot QR + Display</h3>
        <p style="font-size:.85rem;color:var(--text-muted);margin-bottom:.75rem;">Generates QR and projection slide for the practice session (fictional candidates).</p>
        <button class="btn btn-primary" onclick="generatePractice()">Generate Practice QR</button>
        <div id="practice-results" style="display:none;"></div>
      </div>
      <div class="card"><h3>Assisted Mode URLs</h3>
        <p style="font-size:.85rem;color:var(--text-muted);margin-bottom:.75rem;">Assisted-mode ballot URLs for committee members helping delegates vote.</p>
        <div class="form-row">
          <div class="form-group"><label>Election</label>
            <select id="assist-election" style="width:260px;"><option value="">-- Select --</option></select>
          </div>
          <button class="btn btn-primary" onclick="generateAssisted()">Generate Assisted URL</button>
        </div>
        <div id="assist-results" style="display:none;"></div>
      </div>
    </div>

    <!-- LIVE CONTROLS -->
    <div class="section" id="sec-livecontrol">
      <h2>Live Election Controls</h2>
      <div class="card" style="border-left:4px solid var(--warning);"><h3>Environment</h3>
        <div class="form-row">
          <div class="form-group"><label>Target</label>
            <select id="live-env" style="width:200px;"><option value="prod">Production (Convention Day)</option><option value="test" selected>Test Environment</option></select>
          </div>
          <p style="font-size:.8rem;color:var(--danger);align-self:center;"><strong>Caution:</strong> Production controls affect live convention ballots.</p>
        </div>
      </div>
      <div class="card"><h3>1. Select Election &amp; Type</h3>
        <div class="form-row">
          <div class="form-group"><label>Election</label>
            <select id="live-election" style="width:260px;" onchange="updateLivePreview()"><option value="">-- Select --</option></select>
          </div>
          <div class="form-group"><label>Ballot Type</label>
            <select id="live-ballot-type" style="width:220px;" onchange="updateLivePreview()">
              <option value="ranked">Ranked Choice (RCV)</option>
              <option value="runoff">Tiebreaker / Runoff</option>
              <option value="head-to-head">Head-to-Head (2 candidates)</option>
              <option value="slate">Slate Vote (Yes/No)</option>
            </select>
          </div>
        </div>
        <div id="live-type-note" style="font-size:.8rem;color:var(--text-muted);margin-top:.5rem;"></div>
        <div id="live-candidate-preview" style="margin-top:.75rem;"></div>
      </div>
      <div class="card"><h3>2. Push Setup to Google Sheet</h3>
        <p style="font-size:.85rem;color:var(--text-muted);margin-bottom:.75rem;">Creates a Setup tab in the Sheet with election metadata. Do this before opening voting.</p>
        <div class="form-row">
          <div class="form-group"><label>Voting Duration (min)</label><input type="number" id="live-duration" style="width:120px;" min="1"></div>
          <div id="live-duration-hint" style="font-size:.8rem;color:var(--text-muted);align-self:center;"></div>
        </div>
        <div style="display:flex;gap:.75rem;margin-top:.75rem;">
          <button class="btn btn-primary" onclick="pushSetup()">Push Setup to Sheet</button>
          <button class="btn btn-outline" onclick="markCandidatesLive()">Mark Candidates Live</button>
        </div>
        <div id="push-result" style="margin-top:.5rem;"></div>
      </div>
      <div class="card" style="border-left:4px solid var(--success);"><h3>3. Voting Controls</h3>
        <div id="live-status-display" style="margin-bottom:.75rem;"><span class="badge badge-info">Status: Unknown</span></div>
        <div class="form-row">
          <div class="form-group"><label>Floor Count</label><input type="number" id="live-floor-count" style="width:120px;" min="0" value="0"></div>
        </div>
        <div style="display:flex;gap:.75rem;margin-top:.75rem;">
          <button class="btn btn-success" onclick="openVoting()" style="font-size:1rem;padding:.65rem 1.5rem;">&#9654; Open Voting</button>
          <button class="btn btn-danger" onclick="closeVoting()" style="font-size:1rem;padding:.65rem 1.5rem;">&#9632; Close Voting</button>
          <button class="btn btn-outline" onclick="refreshVotingStatus()">Refresh Status</button>
        </div>
        <div id="live-timer-display" style="margin-top:1rem;display:none;">
          <div style="display:flex;align-items:center;gap:1rem;">
            <span style="font-size:2rem;font-weight:700;font-family:monospace;" id="live-timer-clock">00:00</span>
            <span id="live-timer-label" style="color:var(--text-muted);">elapsed</span>
          </div>
          <div style="margin-top:.5rem;">
            <span id="live-vote-count" style="font-size:1.2rem;font-weight:600;color:var(--navy);">0</span>
            <span style="color:var(--text-muted);"> votes received</span>
          </div>
        </div>
      </div>
      <div class="card"><h3>4. Inactive Credentials Report</h3>
        <p style="font-size:.85rem;color:var(--text-muted);margin-bottom:.75rem;">Surrendered and non-issued credentials from local records.</p>
        <div id="inactive-report"></div>
        <button class="btn btn-outline" onclick="refreshInactiveReport()" style="margin-top:.5rem;">Refresh Report</button>
      </div>
    </div>

    <!-- TABULATION -->
    <div class="section" id="sec-tabulate">
      <h2>Run Tabulation</h2>
      <div class="card" style="border-left:4px solid var(--warning);"><h3>Environment &amp; Election</h3>
        <div class="form-row">
          <div class="form-group"><label>Target</label>
            <select id="tab-env" style="width:200px;"><option value="prod">Production</option><option value="test" selected>Test</option></select>
          </div>
          <div class="form-group"><label>Election</label>
            <select id="tab-election" style="width:260px;" onchange="updateTabPreview()"><option value="">-- Select --</option></select>
          </div>
          <button class="btn btn-primary" onclick="fetchBallots()">Fetch Ballots from Sheet</button>
        </div>
        <div id="tab-fetch-status" style="margin-top:.5rem;font-size:.85rem;color:var(--text-muted);"></div>
      </div>
      <div class="card" id="tab-credentials-card" style="display:none;">
        <h3>Credentials Check</h3>
        <div id="tab-credentials-info"></div>
      </div>
      <div class="card" id="tab-run-card" style="display:none;">
        <h3>Run Tabulation</h3>
        <div id="tab-ballot-summary" style="margin-bottom:.75rem;"></div>
        <button class="btn btn-success" onclick="runTabulation()" style="font-size:1rem;padding:.65rem 1.5rem;">Run RCV Tabulation</button>
        <div id="tab-running-status" style="margin-top:.5rem;"></div>
      </div>
      <div class="card" id="tab-results-card" style="display:none;">
        <h3 id="tab-results-title">Results</h3>
        <div id="tab-seat-tabs" style="display:none;margin-bottom:.75rem;"></div>
        <div id="tab-round-nav" class="round-nav" style="display:none;margin-bottom:1rem;">
          <button class="btn btn-sm btn-outline" onclick="tabPrevRound()">&#9664; Prev</button>
          <span id="tab-round-label" style="flex:1;text-align:center;font-weight:600;color:var(--navy);"></span>
          <button class="btn btn-sm btn-outline" onclick="tabNextRound()">Next &#9654;</button>
          <button class="btn btn-sm btn-warning" onclick="tabAutoPlay()" id="tab-auto-btn" style="margin-left:.5rem;">Auto-Play</button>
        </div>
        <div id="tab-round-status" style="margin-bottom:.75rem;display:none;padding:.6rem 1rem;border-radius:6px;font-weight:600;"></div>
        <div id="tab-chart" style="margin-bottom:1rem;"></div>
        <div id="tab-table-wrap" style="max-height:400px;overflow-y:auto;"></div>
      </div>
      <div class="card" id="tab-summary-card" style="display:none;">
        <h3>Final Results Summary</h3>
        <div id="tab-summary-content"></div>
      </div>
    </div>

    <!-- ALL URLs -->
    <div class="section" id="sec-urls">
      <h2>All Generated URLs</h2>
      <div id="all-urls-content"></div>
    </div>
  </div>
</div>

<script>
// ═══════════════════════════════════════════════════════════════════════
//  STATE
// ═══════════════════════════════════════════════════════════════════════
let STATE = {
  elections: {},
  nominees: {},
  delegates: [],
  surrendered: [],
  activeSection: 'dashboard',
  activeElection: null,   // key of election currently being configured in Live Controls
  csvData: null,
  timerInterval: null,
  timerStart: null,
  pollInterval: null,
};

// ═══════════════════════════════════════════════════════════════════════
//  UTILITIES
// ═══════════════════════════════════════════════════════════════════════
function toast(msg, type='success') {
  const c = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = 'toast toast-' + type;
  t.textContent = msg;
  c.appendChild(t);
  setTimeout(() => t.remove(), 4000);
}

async function api(method, path, body) {
  const opts = { method, headers: { 'Content-Type': 'application/json' } };
  if (body) opts.body = JSON.stringify(body);
  try {
    const r = await fetch(path, opts);
    return await r.json();
  } catch (e) {
    toast('Network error: ' + e.message, 'error');
    return { status: 'error', message: e.message };
  }
}

function showSection(name) {
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const sec = document.getElementById('sec-' + name);
  if (sec) sec.classList.add('active');
  const nav = document.querySelector(`[data-section="${name}"]`);
  if (nav) nav.classList.add('active');
  STATE.activeSection = name;
  if (name === 'dashboard') renderDashboard();
  if (name === 'elections') renderElections();
  if (name === 'nominees') renderNominees();
  if (name === 'delegates') renderDelegates();
  if (name === 'surrendered') renderSurrendered();
  if (name === 'generate') populateElectionDropdowns();
  if (name === 'livecontrol') { populateElectionDropdowns(); renderPendingElections(); }
  if (name === 'tabulate') { populateElectionDropdowns(); populateTabDropdown(); }
  if (name === 'urls') renderAllURLs();
}

function copyToClipboard(text) {
  navigator.clipboard.writeText(text).then(() => toast('Copied!')).catch(() => {
    const ta = document.createElement('textarea'); ta.value = text;
    document.body.appendChild(ta); ta.select(); document.execCommand('copy');
    document.body.removeChild(ta); toast('Copied!');
  });
}

function elTypeLabel(t) {
  return { ranked: 'Ranked Choice (RCV)', slate: 'Slate Vote', runoff: 'Runoff / Tiebreaker', 'head-to-head': 'Head-to-Head' }[t] || t;
}

function elStatusBadge(s) {
  const m = { pending: 'badge-info', ready: 'badge-warning', active: 'badge-success', completed: 'badge-info' };
  return `<span class="badge ${m[s] || 'badge-info'}">${(s||'pending').toUpperCase()}</span>`;
}


// ═══════════════════════════════════════════════════════════════════════
//  DATA LOADING
// ═══════════════════════════════════════════════════════════════════════
async function loadAll() {
  const d = await api('GET', '/api/status');
  if (d.status === 'ok') {
    STATE.elections = d.elections || {};
    STATE.nominees = d.nominees || {};
    STATE.delegates = d.delegates || [];
    STATE.surrendered = d.surrendered || [];
  }
  renderDashboard();
}

// ═══════════════════════════════════════════════════════════════════════
//  DASHBOARD
// ═══════════════════════════════════════════════════════════════════════
function renderDashboard() {
  const bar = document.getElementById('status-bar');
  const grid = document.getElementById('dashboard-grid');
  const els = Object.values(STATE.elections);
  const totalNom = Object.values(STATE.nominees).reduce((a,b) => a + b.length, 0);
  bar.innerHTML = `
    <div class="status-item"><span class="status-dot green"></span> ${els.length} Elections</div>
    <div class="status-item"><span class="status-dot ${totalNom ? 'green' : 'yellow'}"></span> ${totalNom} Nominees</div>
    <div class="status-item"><span class="status-dot ${STATE.delegates.length ? 'green' : 'yellow'}"></span> ${STATE.delegates.length} Delegates</div>
    <div class="status-item"><span class="status-dot ${STATE.surrendered.length ? 'red' : 'green'}"></span> ${STATE.surrendered.length} Surrendered</div>
  `;
  grid.innerHTML = els.map(e => {
    const noms = (STATE.nominees[e.key] || []).length;
    const status = e.status || 'pending';
    return `<div class="election-card">
      <h4>${e.label} ${elStatusBadge(status)}</h4>
      <div class="count">${noms}</div>
      <div class="meta">${noms} nominee${noms!==1?'s':''} &middot; ${e.seats} seat${e.seats!==1?'s':''} &middot; ${elTypeLabel(e.type)}</div>
    </div>`;
  }).join('');
}

// ═══════════════════════════════════════════════════════════════════════
//  ELECTIONS CRUD
// ═══════════════════════════════════════════════════════════════════════
function renderElections() {
  const tb = document.getElementById('elections-table');
  const els = Object.values(STATE.elections);
  tb.innerHTML = els.map(e => {
    const noms = (STATE.nominees[e.key] || []).length;
    return `<tr>
      <td><code>${e.key}</code></td>
      <td>${e.label}</td>
      <td>${e.seats}</td>
      <td>${elTypeLabel(e.type)} ${elStatusBadge(e.status || 'pending')}</td>
      <td>${noms}</td>
      <td>
        <button class="btn btn-sm btn-outline" onclick="editElection('${e.key}')">Edit</button>
        <button class="btn btn-sm btn-danger" onclick="deleteElection('${e.key}')">Delete</button>
      </td>
    </tr>`;
  }).join('');
}

function editElection(key) {
  const e = STATE.elections[key];
  if (!e) return;
  document.getElementById('el-key').value = e.key;
  document.getElementById('el-label').value = e.label;
  document.getElementById('el-seats').value = e.seats;
  document.getElementById('el-type').value = e.type;
}

async function saveElection() {
  const key = document.getElementById('el-key').value.trim();
  const label = document.getElementById('el-label').value.trim();
  const seats = parseInt(document.getElementById('el-seats').value) || 1;
  const type = document.getElementById('el-type').value;
  if (!key || !label) { toast('Key and Name are required', 'error'); return; }
  const existing = STATE.elections[key] || {};
  const el = { key, label, seats, type, status: existing.status || 'pending', config: existing.config || {} };
  const r = await api('POST', '/api/elections/save', el);
  if (r.status === 'ok') {
    STATE.elections[key] = el;
    if (!STATE.nominees[key]) STATE.nominees[key] = [];
    toast(`Election "${label}" saved`);
    renderElections();
    document.getElementById('el-key').value = '';
    document.getElementById('el-label').value = '';
    document.getElementById('el-seats').value = '1';
  } else { toast(r.message || 'Error saving', 'error'); }
}

async function deleteElection(key) {
  if (!confirm(`Delete election "${STATE.elections[key]?.label}"?`)) return;
  const r = await api('POST', '/api/elections/delete', { key });
  if (r.status === 'ok') {
    delete STATE.elections[key];
    toast('Election deleted');
    renderElections();
  } else { toast(r.message || 'Error', 'error'); }
}


// ═══════════════════════════════════════════════════════════════════════
//  NOMINEES
// ═══════════════════════════════════════════════════════════════════════
function renderNominees() {
  const tabs = document.getElementById('nominee-tabs');
  const content = document.getElementById('nominee-content');
  const els = Object.values(STATE.elections);
  if (!els.length) { tabs.innerHTML=''; content.innerHTML='<p>No elections defined yet.</p>'; return; }
  const first = STATE.activeElection || els[0].key;
  tabs.innerHTML = els.map(e =>
    `<button class="tab-btn ${e.key===first?'active':''}" onclick="switchNomineeTab('${e.key}')">${e.label}</button>`
  ).join('');
  renderNomineeTab(first);
}

function switchNomineeTab(key) {
  STATE.activeElection = key;
  document.querySelectorAll('#nominee-tabs .tab-btn').forEach(b => b.classList.toggle('active', b.textContent === STATE.elections[key]?.label));
  renderNomineeTab(key);
}

function renderNomineeTab(key) {
  const content = document.getElementById('nominee-content');
  const noms = STATE.nominees[key] || [];
  const e = STATE.elections[key];
  let html = `<div class="card"><h3>Add Nominee to ${e.label}</h3>
    <div class="form-row">
      <div class="form-group"><label>Letter</label><input type="text" id="nom-letter-${key}" value="${nextLetter(noms)}" style="width:60px;" maxlength="1"></div>
      <div class="form-group"><label>Full Name (Last, First)</label><input type="text" id="nom-name-${key}" placeholder="Smith, Jane" style="width:260px;"></div>
      <button class="btn btn-primary" onclick="addNominee('${key}')">Add</button>
    </div></div>`;
  html += `<div class="card"><h3>Nominees (${noms.length})</h3>`;
  if (noms.length) {
    html += `<div class="nominee-list">${noms.map(n =>
      `<span class="nominee-pill"><span class="letter">${n.letter}.</span> ${n.name}
        <button class="remove-btn" onclick="removeNominee('${key}','${n.letter}')" title="Remove">&times;</button>
      </span>`
    ).join('')}</div>`;
  } else { html += '<p style="color:var(--text-muted);">No nominees yet.</p>'; }
  html += '</div>';
  content.innerHTML = html;
}

function nextLetter(noms) {
  if (!noms.length) return 'A';
  const last = noms[noms.length - 1].letter;
  return String.fromCharCode(last.charCodeAt(0) + 1);
}

async function addNominee(key) {
  const letter = document.getElementById('nom-letter-' + key).value.trim().toUpperCase();
  const name = document.getElementById('nom-name-' + key).value.trim();
  if (!letter || !name) { toast('Letter and name required', 'error'); return; }
  const r = await api('POST', '/api/nominees/add', { key, letter, name });
  if (r.status === 'ok') {
    if (!STATE.nominees[key]) STATE.nominees[key] = [];
    STATE.nominees[key].push({ letter, name });
    STATE.nominees[key].sort((a,b) => a.letter.localeCompare(b.letter));
    toast(`Added ${letter}. ${name}`);
    renderNomineeTab(key);
  } else { toast(r.message || 'Error', 'error'); }
}

async function removeNominee(key, letter) {
  const r = await api('POST', '/api/nominees/remove', { key, letter });
  if (r.status === 'ok') {
    STATE.nominees[key] = (STATE.nominees[key]||[]).filter(n => n.letter !== letter);
    toast('Nominee removed');
    renderNomineeTab(key);
  } else { toast(r.message || 'Error', 'error'); }
}


// ═══════════════════════════════════════════════════════════════════════
//  DELEGATES
// ═══════════════════════════════════════════════════════════════════════
function renderDelegates() {
  const tb = document.getElementById('delegate-table');
  const ct = document.getElementById('delegate-count');
  ct.textContent = `Roster (${STATE.delegates.length} loaded)`;
  const q = (document.getElementById('delegate-search')?.value || '').toLowerCase();
  const filtered = STATE.delegates.filter(d => {
    const s = `${d.id} ${d.last} ${d.first} ${d.email||''} ${d.phone||''}`.toLowerCase();
    return s.includes(q);
  });
  tb.innerHTML = filtered.slice(0, 200).map(d =>
    `<tr><td><code>${d.id}</code></td><td>${d.last}, ${d.first}</td><td>${d.email||''}</td><td>${d.phone||''}</td></tr>`
  ).join('');
  if (filtered.length > 200) tb.innerHTML += `<tr><td colspan="4" style="color:var(--text-muted);text-align:center;">... and ${filtered.length-200} more</td></tr>`;
}

function filterDelegates() { renderDelegates(); }

async function addDelegate() {
  const id = document.getElementById('add-del-id').value.trim();
  const last = document.getElementById('add-del-last').value.trim();
  const first = document.getElementById('add-del-first').value.trim();
  const email = document.getElementById('add-del-email').value.trim();
  const phone = document.getElementById('add-del-phone').value.trim();
  if (!id || !last || !first) { toast('ID, Last, and First name required', 'error'); return; }
  const r = await api('POST', '/api/delegates/add', { id, last, first, email, phone });
  if (r.status === 'ok') {
    STATE.delegates.push({ id, last, first, email, phone });
    toast('Delegate added');
    renderDelegates();
    ['add-del-id','add-del-last','add-del-first','add-del-email','add-del-phone'].forEach(x => document.getElementById(x).value = '');
  } else { toast(r.message || 'Error', 'error'); }
}

// ── CSV Import ──
function handleCSVSelect(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const lines = e.target.result.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) { toast('CSV appears empty', 'error'); return; }
    const headers = parseCSVLine(lines[0]);
    STATE.csvData = { headers, lines: lines.slice(1) };
    showCSVMapper(headers, file.name, lines.length - 1);
  };
  reader.readAsText(file);
}

function parseCSVLine(line) {
  const result = []; let current = ''; let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { inQuotes = !inQuotes; }
    else if (ch === ',' && !inQuotes) { result.push(current.trim()); current = ''; }
    else { current += ch; }
  }
  result.push(current.trim());
  return result;
}

function showCSVMapper(headers, filename, rowCount) {
  const mapper = document.getElementById('csv-mapper');
  const fields = [
    { id: 'map-id', label: 'Delegate ID', hint: 'D-6### or A-6###' },
    { id: 'map-last', label: 'Last Name' },
    { id: 'map-first', label: 'First Name' },
    { id: 'map-email', label: 'Email', optional: true },
    { id: 'map-phone', label: 'Phone', optional: true },
  ];
  const options = headers.map((h,i) => `<option value="${i}">${h}</option>`).join('');
  const guesses = guessColumnMapping(headers);
  let html = `<div style="margin-bottom:.75rem;"><strong>${filename}</strong> — ${rowCount} rows, ${headers.length} columns</div>`;
  html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:.5rem .75rem;max-width:500px;">';
  fields.forEach(f => {
    const sel = guesses[f.id] !== undefined ? guesses[f.id] : '';
    html += `<label style="font-size:.85rem;font-weight:600;">${f.label}${f.optional?' (opt)':''}</label>`;
    html += `<select id="${f.id}" style="font-size:.85rem;"><option value="">-- skip --</option>${options}</select>`;
  });
  html += '</div>';
  html += `<button class="btn btn-success" onclick="importCSV()" style="margin-top:.75rem;">Import ${rowCount} Delegates</button>`;
  mapper.innerHTML = html;
  mapper.style.display = 'block';
  // Apply guesses
  Object.entries(guesses).forEach(([id, idx]) => {
    const sel = document.getElementById(id);
    if (sel && idx !== undefined) sel.value = idx;
  });
}

function guessColumnMapping(headers) {
  const g = {};
  const lower = headers.map(h => h.toLowerCase().replace(/[^a-z0-9]/g,''));
  const find = (patterns) => lower.findIndex(h => patterns.some(p => h.includes(p)));
  const idIdx = find(['delegateid','delid','credentialid','id','number','num','badge']);
  const lastIdx = find(['lastname','last','surname','family']);
  const firstIdx = find(['firstname','first','given']);
  const emailIdx = find(['email','mail']);
  const phoneIdx = find(['phone','mobile','cell','tel']);
  if (idIdx >= 0) g['map-id'] = idIdx;
  if (lastIdx >= 0) g['map-last'] = lastIdx;
  if (firstIdx >= 0) g['map-first'] = firstIdx;
  if (emailIdx >= 0) g['map-email'] = emailIdx;
  if (phoneIdx >= 0) g['map-phone'] = phoneIdx;
  return g;
}

async function importCSV() {
  const getCol = id => { const s = document.getElementById(id); return s ? parseInt(s.value) : -1; };
  const idCol = getCol('map-id'), lastCol = getCol('map-last'), firstCol = getCol('map-first');
  const emailCol = getCol('map-email'), phoneCol = getCol('map-phone');
  if (isNaN(idCol) || idCol < 0 || isNaN(lastCol) || lastCol < 0 || isNaN(firstCol) || firstCol < 0) {
    toast('ID, Last Name, and First Name columns are required', 'error'); return;
  }
  const rows = STATE.csvData.lines.map(line => {
    const cols = parseCSVLine(line);
    return {
      id: (cols[idCol]||'').trim(), last: (cols[lastCol]||'').trim(), first: (cols[firstCol]||'').trim(),
      email: emailCol >= 0 ? (cols[emailCol]||'').trim() : '', phone: phoneCol >= 0 ? (cols[phoneCol]||'').trim() : ''
    };
  }).filter(r => r.id && r.last);
  const r = await api('POST', '/api/delegates/import', { delegates: rows });
  if (r.status === 'ok') {
    STATE.delegates = r.delegates || rows;
    toast(`Imported ${rows.length} delegates`);
    document.getElementById('csv-mapper').style.display = 'none';
    renderDelegates();
  } else { toast(r.message || 'Error', 'error'); }
}


// ═══════════════════════════════════════════════════════════════════════
//  SURRENDERED CREDENTIALS
// ═══════════════════════════════════════════════════════════════════════
function renderSurrendered() {
  const tb = document.getElementById('surr-table');
  const ct = document.getElementById('surr-count');
  ct.textContent = `Surrendered List (${STATE.surrendered.length})`;
  tb.innerHTML = STATE.surrendered.map(s => {
    const del = STATE.delegates.find(d => d.id === s.id);
    const name = del ? `${del.last}, ${del.first}` : '';
    return `<tr><td><code>${s.id}</code></td><td>${name}</td><td>${s.status}</td><td>${s.time||''}</td><td>${s.note||''}</td>
      <td><button class="btn btn-sm btn-outline" onclick="removeSurrender('${s.id}')" title="Remove">&times;</button></td></tr>`;
  }).join('');
}

async function recordSurrender() {
  const id = document.getElementById('surr-id').value.trim();
  const status = document.getElementById('surr-status').value;
  const time = document.getElementById('surr-time').value.trim() || new Date().toLocaleString('en-US', {timeZone:'America/Chicago'});
  const note = document.getElementById('surr-note').value.trim();
  if (!id) { toast('Delegate ID required', 'error'); return; }
  const r = await api('POST', '/api/surrendered/add', { id, status, time, note });
  if (r.status === 'ok') {
    STATE.surrendered.push({ id, status, time, note });
    toast('Surrender recorded');
    renderSurrendered();
    ['surr-id','surr-time','surr-note'].forEach(x => document.getElementById(x).value = '');
  } else { toast(r.message || 'Error', 'error'); }
}

async function removeSurrender(id) {
  const r = await api('POST', '/api/surrendered/remove', { id });
  if (r.status === 'ok') {
    STATE.surrendered = STATE.surrendered.filter(s => s.id !== id);
    toast('Removed'); renderSurrendered();
  }
}

// ═══════════════════════════════════════════════════════════════════════
//  QR / DISPLAY GENERATION
// ═══════════════════════════════════════════════════════════════════════
function populateElectionDropdowns() {
  const els = Object.values(STATE.elections);
  const opts = els.map(e => `<option value="${e.key}">${e.label}</option>`).join('');
  ['gen-election','assist-election','live-election'].forEach(id => {
    const sel = document.getElementById(id);
    if (sel) { const cur = sel.value; sel.innerHTML = '<option value="">-- Select --</option>' + opts; if (cur) sel.value = cur; }
  });
}

function toggleTimerOptions() {
  const show = document.getElementById('gen-timer').value === 'yes';
  document.getElementById('gen-duration-group').style.display = show ? '' : 'none';
  if (show) updateTimerSuggestion('gen');
}

function updateTimerSuggestion(prefix) {
  const elKey = document.getElementById(prefix === 'gen' ? 'gen-election' : 'live-election').value;
  if (!elKey) return;
  const noms = (STATE.nominees[elKey] || []).length;
  const suggested = (noms * 2) + 34;
  const input = document.getElementById(prefix === 'gen' ? 'gen-duration' : 'live-duration');
  if (input && !input.value) input.value = suggested;
  const hint = document.getElementById(prefix === 'gen' ? 'gen-duration-group' : 'live-duration-hint');
  if (hint && prefix === 'live') hint.textContent = `Suggested: ${suggested} min (${noms} candidates × 2 + 34)`;
}

function updateGenPreview() {
  const key = document.getElementById('gen-election').value;
  const div = document.getElementById('gen-candidate-preview');
  if (!key) { div.innerHTML = ''; return; }
  const noms = STATE.nominees[key] || [];
  div.innerHTML = noms.length ? `Candidates: ${noms.map(n => `${n.letter}. ${n.name}`).join(', ')}` : 'No nominees assigned yet.';
  updateTimerSuggestion('gen');
}

async function generateQR() {
  const key = document.getElementById('gen-election').value;
  if (!key) { toast('Select an election', 'error'); return; }
  const noms = STATE.nominees[key] || [];
  if (!noms.length) { toast('No nominees for this election', 'error'); return; }
  const timerOn = document.getElementById('gen-timer').value === 'yes';
  const duration = timerOn ? parseInt(document.getElementById('gen-duration').value) || null : null;
  const r = await api('POST', '/api/generate', { key, duration });
  if (r.status === 'ok') {
    const div = document.getElementById('gen-results');
    div.style.display = 'block';
    div.innerHTML = `<div class="results-panel"><h4>Generated Files</h4>
      ${r.qr ? `<div class="qr-preview"><img src="${r.qr}" alt="QR Code"><div>
        <div class="url-display" onclick="copyToClipboard('${r.ballot_url}')" title="Click to copy">${r.ballot_url}</div>
        ${r.admin_url ? `<div class="url-display" onclick="copyToClipboard('${r.admin_url}')" title="Click to copy">Admin: ${r.admin_url}</div>` : ''}
        <div class="file-link">QR PNG: ${r.qr_file || 'saved'}</div>
        <div class="file-link">Display HTML: ${r.display_file || 'saved'}</div>
      </div></div>` : '<p>QR library not available</p>'}
    </div>`;
    toast('QR code and display generated');
  } else { toast(r.message || 'Generation failed', 'error'); }
}

async function generatePractice() {
  const r = await api('POST', '/api/generate-practice', {});
  if (r.status === 'ok') {
    const div = document.getElementById('practice-results');
    div.style.display = 'block';
    div.innerHTML = `<div class="results-panel"><h4>Practice Files Generated</h4>
      ${r.qr ? `<div class="qr-preview"><img src="${r.qr}" alt="Practice QR"><div>
        <div class="url-display" onclick="copyToClipboard('${r.ballot_url}')" title="Click to copy">${r.ballot_url}</div>
      </div></div>` : ''}
    </div>`;
    toast('Practice QR generated');
  } else { toast(r.message || 'Error', 'error'); }
}

async function generateAssisted() {
  const key = document.getElementById('assist-election').value;
  if (!key) { toast('Select an election', 'error'); return; }
  const noms = STATE.nominees[key] || [];
  if (!noms.length) { toast('No nominees', 'error'); return; }
  const r = await api('POST', '/api/generate-assisted', { key });
  if (r.status === 'ok') {
    const div = document.getElementById('assist-results');
    div.style.display = 'block';
    div.innerHTML = `<div class="results-panel"><h4>Assisted Mode URL</h4>
      <div class="url-display" onclick="copyToClipboard('${r.url}')" title="Click to copy">${r.url}</div>
    </div>`;
    toast('Assisted URL generated');
  } else { toast(r.message || 'Error', 'error'); }
}


// ═══════════════════════════════════════════════════════════════════════
//  LIVE CONTROLS — SAVE & SWITCH ELECTIONS
// ═══════════════════════════════════════════════════════════════════════

function renderPendingElections() {
  // Build a status panel showing all elections so you can switch between them
  const els = Object.values(STATE.elections);
  const liveEl = document.getElementById('live-election').value;
  let html = '<div style="display:flex;flex-wrap:wrap;gap:.5rem;margin-bottom:1rem;">';
  els.forEach(e => {
    const status = e.status || 'pending';
    const isCurrent = e.key === liveEl;
    const colors = {
      pending: 'background:#e3f2fd;border-color:#90caf9;',
      ready: 'background:#fff8e1;border-color:#ffe082;',
      active: 'background:#e8f5e9;border-color:#a5d6a7;',
      completed: 'background:#f3e5f5;border-color:#ce93d8;'
    };
    html += `<div onclick="switchToElection('${e.key}')" style="cursor:pointer;padding:.5rem .85rem;border-radius:8px;border:2px solid;${colors[status]||colors.pending}${isCurrent?'box-shadow:0 0 0 2px var(--navy);':''}font-size:.85rem;">
      <strong>${e.label}</strong><br>
      <span style="font-size:.75rem;">${status.toUpperCase()}${(e.config?.duration) ? ' · '+e.config.duration+'m' : ''}${(STATE.nominees[e.key]||[]).length ? ' · '+(STATE.nominees[e.key]||[]).length+' nom' : ''}</span>
    </div>`;
  });
  html += '</div>';
  // Insert before the first card in live controls
  let container = document.getElementById('pending-elections-bar');
  if (!container) {
    container = document.createElement('div');
    container.id = 'pending-elections-bar';
    const secLive = document.getElementById('sec-livecontrol');
    secLive.insertBefore(container, secLive.querySelector('.card'));
  }
  container.innerHTML = '<h3 style="margin-bottom:.5rem;font-size:.95rem;color:var(--navy);">Election Queue — click to switch</h3>' + html;
}

function switchToElection(key) {
  // Auto-save current election config before switching
  const currentKey = document.getElementById('live-election').value;
  if (currentKey && currentKey !== key) {
    saveElectionConfig(currentKey);
  }
  // Load the target election
  document.getElementById('live-election').value = key;
  const e = STATE.elections[key];
  if (e) {
    // Restore saved config
    if (e.config?.ballotType) document.getElementById('live-ballot-type').value = e.config.ballotType;
    if (e.config?.duration) document.getElementById('live-duration').value = e.config.duration;
    if (e.config?.floorCount !== undefined) document.getElementById('live-floor-count').value = e.config.floorCount;
    if (e.type) document.getElementById('live-ballot-type').value = e.config?.ballotType || e.type;
  }
  updateLivePreview();
  renderPendingElections();
  toast(`Switched to: ${e?.label || key}`);
}

function saveElectionConfig(key) {
  if (!key) return;
  const e = STATE.elections[key];
  if (!e) return;
  e.config = e.config || {};
  e.config.ballotType = document.getElementById('live-ballot-type').value;
  e.config.duration = parseInt(document.getElementById('live-duration').value) || null;
  e.config.floorCount = parseInt(document.getElementById('live-floor-count').value) || 0;
  // Persist to server
  api('POST', '/api/elections/save', e);
}

async function saveCurrentElectionConfig() {
  const key = document.getElementById('live-election').value;
  if (!key) { toast('No election selected', 'error'); return; }
  saveElectionConfig(key);
  // Update status to "ready" if still pending
  const e = STATE.elections[key];
  if (e && e.status === 'pending' && (STATE.nominees[key]||[]).length > 0) {
    e.status = 'ready';
    await api('POST', '/api/elections/save', e);
  }
  toast(`Configuration saved for ${e?.label || key}`);
  renderPendingElections();
}

// ═══════════════════════════════════════════════════════════════════════
//  LIVE CONTROLS — VOTING
// ═══════════════════════════════════════════════════════════════════════
function updateLivePreview() {
  const key = document.getElementById('live-election').value;
  const type = document.getElementById('live-ballot-type').value;
  const noteDiv = document.getElementById('live-type-note');
  const previewDiv = document.getElementById('live-candidate-preview');
  const typeNotes = {
    ranked: 'Delegates rank candidates in order of preference. Eliminated candidates\' votes transfer to next choice.',
    runoff: 'Top candidates from a prior round. Delegates pick one.',
    'head-to-head': 'Two candidates only. Simple majority wins.',
    slate: 'Full slate presented for Yes/No vote. No individual ranking.'
  };
  noteDiv.textContent = typeNotes[type] || '';
  if (key) {
    const noms = STATE.nominees[key] || [];
    previewDiv.innerHTML = noms.length
      ? `<div class="nominee-list">${noms.map(n => `<span class="nominee-pill"><span class="letter">${n.letter}.</span> ${n.name}</span>`).join('')}</div>`
      : '<p style="color:var(--warning);">No nominees assigned. Add nominees first.</p>';
    updateTimerSuggestion('live');
  } else { previewDiv.innerHTML = ''; }
}

async function pushSetup() {
  const key = document.getElementById('live-election').value;
  if (!key) { toast('Select an election first', 'error'); return; }
  saveElectionConfig(key);
  const env = document.getElementById('live-env').value;
  const duration = parseInt(document.getElementById('live-duration').value) || null;
  const noms = STATE.nominees[key] || [];
  const ballotType = document.getElementById('live-ballot-type').value;
  const floorCount = parseInt(document.getElementById('live-floor-count').value) || 0;
  const r = await api('POST', '/api/push-setup', { key, env, duration, ballotType, floorCount, nominees: noms });
  const div = document.getElementById('push-result');
  if (r.status === 'ok') {
    div.innerHTML = `<span style="color:var(--success);">Setup pushed to ${env} Sheet.</span>`;
    toast('Setup pushed to Google Sheet');
  } else {
    div.innerHTML = `<span style="color:var(--danger);">${r.message||'Error pushing setup'}</span>`;
    toast(r.message || 'Push failed', 'error');
  }
}

async function markCandidatesLive() {
  const key = document.getElementById('live-election').value;
  if (!key) { toast('Select an election', 'error'); return; }
  const env = document.getElementById('live-env').value;
  const r = await api('POST', '/api/mark-live', { key, env });
  if (r.status === 'ok') { toast('Candidates marked as live'); }
  else { toast(r.message || 'Error', 'error'); }
}

async function openVoting() {
  const key = document.getElementById('live-election').value;
  if (!key) { toast('Select an election', 'error'); return; }
  saveElectionConfig(key);
  const env = document.getElementById('live-env').value;
  const floorCount = parseInt(document.getElementById('live-floor-count').value) || 0;
  const r = await api('POST', '/api/voting/open', { key, env, floorCount });
  if (r.status === 'ok') {
    toast('Voting opened!');
    // Update status to active
    if (STATE.elections[key]) { STATE.elections[key].status = 'active'; api('POST', '/api/elections/save', STATE.elections[key]); }
    startTimer();
    startPolling(key, env);
    document.getElementById('live-status-display').innerHTML = '<span class="badge badge-success">VOTING OPEN</span>';
    renderPendingElections();
  } else { toast(r.message || 'Failed to open voting', 'error'); }
}

async function closeVoting() {
  const key = document.getElementById('live-election').value;
  if (!key) { toast('Select an election', 'error'); return; }
  if (!confirm('Close voting for ' + (STATE.elections[key]?.label || key) + '?')) return;
  const env = document.getElementById('live-env').value;
  const r = await api('POST', '/api/voting/close', { key, env });
  if (r.status === 'ok') {
    toast('Voting closed');
    if (STATE.elections[key]) { STATE.elections[key].status = 'completed'; api('POST', '/api/elections/save', STATE.elections[key]); }
    stopTimer(); stopPolling();
    document.getElementById('live-status-display').innerHTML = '<span class="badge badge-info">VOTING CLOSED</span>';
    renderPendingElections();
  } else { toast(r.message || 'Failed to close voting', 'error'); }
}

async function refreshVotingStatus() {
  const key = document.getElementById('live-election').value;
  if (!key) return;
  const env = document.getElementById('live-env').value;
  const r = await api('POST', '/api/voting/status', { key, env });
  if (r.status === 'ok') {
    const d = r.data || {};
    const isOpen = d.accepting === true || d.accepting === 'true';
    document.getElementById('live-status-display').innerHTML = isOpen
      ? '<span class="badge badge-success">VOTING OPEN</span>'
      : '<span class="badge badge-info">VOTING CLOSED</span>';
    if (d.count !== undefined) document.getElementById('live-vote-count').textContent = d.count;
  }
}


// ═══════════════════════════════════════════════════════════════════════
//  TIMER & POLLING
// ═══════════════════════════════════════════════════════════════════════
function startTimer() {
  stopTimer();
  STATE.timerStart = Date.now();
  const display = document.getElementById('live-timer-display');
  display.style.display = 'block';
  STATE.timerInterval = setInterval(() => {
    const elapsed = Math.floor((Date.now() - STATE.timerStart) / 1000);
    const m = String(Math.floor(elapsed / 60)).padStart(2, '0');
    const s = String(elapsed % 60).padStart(2, '0');
    document.getElementById('live-timer-clock').textContent = `${m}:${s}`;
  }, 1000);
}

function stopTimer() {
  if (STATE.timerInterval) { clearInterval(STATE.timerInterval); STATE.timerInterval = null; }
}

function startPolling(key, env) {
  stopPolling();
  STATE.pollInterval = setInterval(async () => {
    const r = await api('POST', '/api/voting/status', { key, env });
    if (r.status === 'ok' && r.data) {
      if (r.data.count !== undefined) document.getElementById('live-vote-count').textContent = r.data.count;
    }
  }, 10000);
}

function stopPolling() {
  if (STATE.pollInterval) { clearInterval(STATE.pollInterval); STATE.pollInterval = null; }
}

// ═══════════════════════════════════════════════════════════════════════
//  INACTIVE CREDENTIALS REPORT
// ═══════════════════════════════════════════════════════════════════════
function refreshInactiveReport() {
  const div = document.getElementById('inactive-report');
  const surr = STATE.surrendered;
  if (!surr.length) { div.innerHTML = '<p style="color:var(--text-muted);">No surrendered credentials on file.</p>'; return; }
  let html = `<table><thead><tr><th>ID</th><th>Name</th><th>Status</th><th>Time</th><th>Note</th></tr></thead><tbody>`;
  surr.forEach(s => {
    const del = STATE.delegates.find(d => d.id === s.id);
    const name = del ? `${del.last}, ${del.first}` : '';
    html += `<tr><td><code>${s.id}</code></td><td>${name}</td><td>${s.status}</td><td>${s.time||''}</td><td>${s.note||''}</td></tr>`;
  });
  html += '</tbody></table>';
  html += `<p style="margin-top:.5rem;font-size:.85rem;color:var(--text-muted);">Total: ${surr.length} credential${surr.length!==1?'s':''}</p>`;
  div.innerHTML = html;
}

// ═══════════════════════════════════════════════════════════════════════
//  ALL URLs
// ═══════════════════════════════════════════════════════════════════════
function renderAllURLs() {
  const div = document.getElementById('all-urls-content');
  const els = Object.values(STATE.elections);
  if (!els.length) { div.innerHTML = '<p>No elections configured.</p>'; return; }
  let html = '';
  els.forEach(e => {
    const noms = STATE.nominees[e.key] || [];
    html += `<div class="card"><h3>${e.label}</h3>`;
    if (noms.length) {
      const cp = noms.map(n => `${n.letter}-${encodeURIComponent(n.name)}`).join('-');
      const ballotUrl = `${location.protocol}//${location.host}` !== 'null' ? `https://jnmpub25.github.io/IA3CD2026DC/ballot.html?election=${e.key}&candidates=${cp}&seats=${e.seats}` : '';
      html += `<div class="url-display" onclick="copyToClipboard(this.textContent)" title="Click to copy">Ballot: ${ballotUrl}</div>`;
      html += `<div class="url-display" onclick="copyToClipboard(this.textContent)" title="Click to copy">Admin: ${ballotUrl}&admin=1</div>`;
      html += `<div class="url-display" onclick="copyToClipboard(this.textContent)" title="Click to copy">Assisted: ${ballotUrl}&mode=assisted</div>`;
    } else {
      html += '<p style="color:var(--text-muted);">No nominees — URLs not available yet.</p>';
    }
    html += '</div>';
  });
  div.innerHTML = html;
}

// ═══════════════════════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  loadAll().then(() => {
    populateElectionDropdowns();
  });
  // CSV drag-and-drop
  const drop = document.getElementById('csv-drop');
  if (drop) {
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag-over'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
    drop.addEventListener('drop', e => {
      e.preventDefault(); drop.classList.remove('drag-over');
      const file = e.dataTransfer.files[0];
      if (file) { document.getElementById('csv-file').files = e.dataTransfer.files; handleCSVSelect(document.getElementById('csv-file')); }
    });
  }
});

// ═══════════════════════════════════════════════════════════════════════
//  TABULATION ENGINE (RCV / IRV)
// ═══════════════════════════════════════════════════════════════════════

const MAJORITY_THRESHOLD = 0.5;
const ELIM_THRESHOLD = 0.15;
const SCC_FIRST_BALLOT_MAX = 0.5;

let tabState = {
  ballots: [],         // parsed Ballot objects
  spoiled: [],         // flagged ballots (surrendered/non-issued)
  candidateKeys: [],
  candidateMap: {},    // letter → name
  allRounds: [],
  electedAll: [],
  currentRoundIdx: 0,
  autoInterval: null,
  seats: 1,
  electionKey: '',
};

function populateTabDropdown() {
  const sel = document.getElementById('tab-election');
  if (!sel) return;
  const cur = sel.value;
  const els = Object.values(STATE.elections);
  sel.innerHTML = '<option value="">-- Select --</option>' + els.map(e =>
    `<option value="${e.key}">${e.label} (${e.seats} seat${e.seats!==1?'s':''})</option>`
  ).join('');
  if (cur) sel.value = cur;
}

function updateTabPreview() {
  const key = document.getElementById('tab-election').value;
  const status = document.getElementById('tab-fetch-status');
  if (!key) { status.innerHTML = ''; return; }
  const e = STATE.elections[key];
  const noms = STATE.nominees[key] || [];
  status.innerHTML = `${e.label} &middot; ${e.seats} seat${e.seats!==1?'s':''} &middot; ${noms.length} nominees &middot; Type: ${elTypeLabel(e.type)}`;
}

async function fetchBallots() {
  const key = document.getElementById('tab-election').value;
  const env = document.getElementById('tab-env').value;
  if (!key) { toast('Select an election', 'error'); return; }
  const status = document.getElementById('tab-fetch-status');
  status.innerHTML = '<span style="color:var(--navy);">Fetching ballots from Google Sheet...</span>';

  const r = await api('POST', '/api/export-ballots', { key, env });
  if (r.status !== 'ok') { status.innerHTML = `<span style="color:var(--danger);">${r.message || 'Error fetching'}</span>`; return; }

  const rows = r.rows || [];
  status.innerHTML = `<span style="color:var(--success);">Fetched ${rows.length} ballot row${rows.length!==1?'s':''}.</span>`;

  // Parse ballots
  const noms = STATE.nominees[key] || [];
  tabState.candidateKeys = noms.map(n => n.letter);
  tabState.candidateMap = {};
  noms.forEach(n => { tabState.candidateMap[n.letter] = n.name; });
  tabState.electionKey = key;
  tabState.seats = STATE.elections[key]?.seats || 1;

  // Detect format and parse rankings
  const parsed = [];
  for (const row of rows) {
    if ((row['Is Test'] || '').toUpperCase() === 'YES') continue;
    const delNum = row['Delegate Number'] || row['Delegate_Number'] || '';
    if (!delNum) continue;
    const rankings = {};
    let foundA = false;
    for (const k of tabState.candidateKeys) {
      const col = `Candidate ${k} Rank`;
      if (row[col] !== undefined) {
        foundA = true;
        const v = parseInt(row[col]);
        rankings[k] = isNaN(v) ? 0 : v;
      }
    }
    if (!foundA) {
      const ordinals = ['1st','2nd','3rd','4th','5th','6th','7th','8th','9th','10th','11th','12th','13th','14th','15th','16th','17th'];
      for (let pos = 0; pos < ordinals.length; pos++) {
        const val = (row[ordinals[pos]] || '').trim().toUpperCase();
        if (val && val.length === 1 && val.match(/[A-Z]/)) { rankings[val] = pos + 1; }
      }
    }
    parsed.push({ delegateNumber: delNum.trim(), rankings, timestamp: row['Timestamp'] || '' });
  }

  // Credentials check
  const surrendered = STATE.surrendered || [];
  tabState.ballots = [];
  tabState.spoiled = [];
  for (const b of parsed) {
    const match = surrendered.find(s => s.id === b.delegateNumber);
    if (match) {
      tabState.spoiled.push({ ...b, reason: match.status === 'non-issued' ? 'Non-Issued Credential' : 'Surrendered Credential', reasonType: match.status });
    } else {
      tabState.ballots.push(b);
    }
  }

  // Show credentials card
  const credCard = document.getElementById('tab-credentials-card');
  const credInfo = document.getElementById('tab-credentials-info');
  credCard.style.display = 'block';
  let credHtml = `<div style="display:flex;gap:1rem;flex-wrap:wrap;margin-bottom:.75rem;">
    <span class="badge badge-info">${parsed.length} ballots loaded</span>
    <span class="badge badge-success">${tabState.ballots.length} valid</span>`;
  if (tabState.spoiled.length > 0) {
    const nSurr = tabState.spoiled.filter(s => s.reasonType !== 'non-issued').length;
    const nNI = tabState.spoiled.filter(s => s.reasonType === 'non-issued').length;
    credHtml += `<span class="badge badge-warning">${tabState.spoiled.length} flagged</span>`;
    if (nSurr) credHtml += `<span class="badge" style="background:#fff3cd;color:#856404;">${nSurr} surrendered</span>`;
    if (nNI) credHtml += `<span class="badge" style="background:#fde8e8;color:#c0392b;">${nNI} non-issued</span>`;
  }
  credHtml += '</div>';
  if (surrendered.length > 0) {
    credHtml += '<details style="margin-top:.5rem;"><summary style="cursor:pointer;font-weight:600;font-size:.85rem;color:var(--navy);">Full Inactive Credentials Report (' + surrendered.length + ')</summary>';
    credHtml += '<table style="margin-top:.5rem;"><thead><tr><th>ID</th><th>Name</th><th>Status</th><th>Voted?</th><th>Note</th></tr></thead><tbody>';
    surrendered.forEach(s => {
      const del = STATE.delegates.find(d => d.id === s.id);
      const name = del ? `${del.last}, ${del.first}` : '';
      const voted = parsed.some(b => b.delegateNumber === s.id);
      credHtml += `<tr><td><code>${s.id}</code></td><td>${name}</td><td>${s.status}</td>
        <td>${voted ? '<span style="color:var(--danger);font-weight:600;">YES - FLAGGED</span>' : 'No'}</td>
        <td>${s.note||''}</td></tr>`;
    });
    credHtml += '</tbody></table></details>';
  }
  if (tabState.spoiled.length > 0) {
    credHtml += '<details style="margin-top:.5rem;" open><summary style="cursor:pointer;font-weight:600;font-size:.85rem;color:var(--danger);">Flagged Ballots (' + tabState.spoiled.length + ')</summary>';
    credHtml += '<table style="margin-top:.5rem;"><thead><tr><th>Delegate</th><th>Reason</th></tr></thead><tbody>';
    tabState.spoiled.forEach(s => {
      credHtml += `<tr><td><code>${s.delegateNumber}</code></td><td>${s.reason}</td></tr>`;
    });
    credHtml += '</tbody></table></details>';
  }
  credInfo.innerHTML = credHtml;

  // Show run card
  document.getElementById('tab-run-card').style.display = 'block';
  document.getElementById('tab-ballot-summary').innerHTML =
    `<strong>${tabState.ballots.length}</strong> valid ballots ready. ` +
    `<strong>${tabState.candidateKeys.length}</strong> candidates: ${noms.map(n => `${n.letter}. ${n.name}`).join(', ')}`;
}

// ── IRV Core Logic ──
function tabCountVotes(ballots, active) {
  const counts = {}; active.forEach(k => counts[k] = 0);
  let exhausted = 0;
  for (const b of ballots) {
    let bestKey = null, bestRank = null;
    for (const k of active) {
      const r = b.rankings[k];
      if (r && r > 0 && (bestRank === null || r < bestRank)) { bestRank = r; bestKey = k; }
    }
    if (bestKey) counts[bestKey]++;
    else exhausted++;
  }
  const totalActive = Object.values(counts).reduce((a,b) => a+b, 0);
  return { voteCounts: counts, totalActive, exhausted };
}

function runIRVTabulation() {
  const { ballots, candidateKeys, seats } = tabState;
  tabState.allRounds = [];
  tabState.electedAll = [];
  const maxSeat1R1Wins = Math.max(1, Math.floor(seats * SCC_FIRST_BALLOT_MAX));

  for (let seat = 1; seat <= seats; seat++) {
    let active = candidateKeys.filter(k => !tabState.electedAll.includes(k));
    if (active.length === 0) break;
    if (active.length === 1) {
      const w = active[0];
      const { voteCounts, totalActive, exhausted } = tabCountVotes(ballots, active);
      tabState.electedAll.push(w);
      tabState.allRounds.push({ seat, seatRound: 1, voteCounts, totalActive, exhausted,
        majorityNeeded: Math.floor(totalActive * MAJORITY_THRESHOLD) + 1,
        active: [...active], elected: [w], eliminated: [], elimReason: 'Uncontested' });
      continue;
    }
    let seatActive = [...active];
    let seatRound = 1;
    let seatDone = false;
    while (seatActive.length > 1 && !seatDone) {
      const { voteCounts, totalActive, exhausted } = tabCountVotes(ballots, seatActive);
      const majorityNeeded = Math.floor(totalActive * MAJORITY_THRESHOLD) + 1;
      const thresholdVotes = totalActive * ELIM_THRESHOLD;
      const winners = seatActive.filter(k => voteCounts[k] >= majorityNeeded);
      if (winners.length > 0) {
        tabState.electedAll.push(winners[0]);
        tabState.allRounds.push({ seat, seatRound, voteCounts, totalActive, exhausted, majorityNeeded,
          active: [...seatActive], elected: [winners[0]], eliminated: [], elimReason: '' });
        seatDone = true; break;
      }
      let toEliminate = []; let elimReason = '';
      const sorted = seatActive.slice().sort((a,b) => voteCounts[a] - voteCounts[b]);
      if (seatRound === 1) {
        const below = seatActive.filter(k => voteCounts[k] < thresholdVotes);
        if (below.length > 0 && below.length < seatActive.length) {
          toEliminate = below;
          elimReason = `Below 15% threshold (${Math.ceil(thresholdVotes)} votes needed)`;
        }
      }
      if (toEliminate.length === 0) {
        const lowest = voteCounts[sorted[0]];
        const tied = sorted.filter(k => voteCounts[k] === lowest);
        if (tied.length >= seatActive.length) {
      const sorted = seatActive.slice().sort((a,b) => voteCounts[a] - voteCounts[b]);
      if (seatRound === 1) {
        const below = seatActive.filter(k => voteCounts[k] < thresholdVotes);
        if (below.length > 0 && below.length < seatActive.length) {
          toEliminate = below;
          elimReason = `Below 15% threshold (${Math.ceil(thresholdVotes)} votes needed)`;
        }
      }
      if (toEliminate.length === 0) {
        const lowest = voteCounts[sorted[0]];
        const tied = sorted.filter(k => voteCounts[k] === lowest);
        if (tied.length >= seatActive.length) {
          tabState.allRounds.push({ seat, seatRound, voteCounts, totalActive, exhausted, majorityNeeded,
            active: [...seatActive], elected: [], eliminated: [], elimReason: 'All tied - no winner' });
          seatDone = true; break;
        }
        toEliminate = tied;
        elimReason = `Lowest vote count (${lowest} vote${lowest!==1?'s':''})`;
      }
      tabState.allRounds.push({ seat, seatRound, voteCounts, totalActive, exhausted, majorityNeeded,
        active: [...seatActive], elected: [], eliminated: [...toEliminate], elimReason });
      seatActive = seatActive.filter(k => !toEliminate.includes(k));
      seatRound++;
      if (seatActive.length === 1) {
        const w = seatActive[0];
        const r2 = tabCountVotes(ballots, seatActive);
        tabState.electedAll.push(w);
        tabState.allRounds.push({ seat, seatRound, voteCounts: r2.voteCounts, totalActive: r2.totalActive,
          exhausted: r2.exhausted, majorityNeeded: Math.floor(r2.totalActive * MAJORITY_THRESHOLD) + 1,
          active: [...seatActive], elected: [w], eliminated: [], elimReason: 'Last candidate standing' });
        seatDone = true;
      }
    }
  }
}

function runTabulation() {
  if (tabState.ballots.length === 0) { toast('No ballots loaded', 'error'); return; }
  document.getElementById('tab-running-status').innerHTML = '<span style="color:var(--navy);">Running tabulation...</span>';
  setTimeout(() => {
    runIRVTabulation();
    tabState.currentRoundIdx = 0;
    document.getElementById('tab-results-card').style.display = 'block';
    const e = STATE.elections[tabState.electionKey];
    document.getElementById('tab-results-title').textContent = `${e?.label || tabState.electionKey} - Round-by-Round Results`;
    const tabsDiv = document.getElementById('tab-seat-tabs');
    if (tabState.seats > 1) {
      tabsDiv.style.display = 'flex'; tabsDiv.style.gap = '.5rem'; tabsDiv.style.flexWrap = 'wrap';
      const seats = [...new Set(tabState.allRounds.map(r => r.seat))];
      tabsDiv.innerHTML = seats.map(s => {
        const winner = tabState.electedAll[s-1];
        const winName = winner ? tabState.candidateMap[winner] : '';
        return `<button class="btn btn-sm btn-outline" onclick="tabGoToSeat(${s})" id="tab-seat-btn-${s}">Seat ${s}${winName ? ': '+winName : ''}</button>`;
      }).join('');
    } else { tabsDiv.style.display = 'none'; }
    document.getElementById('tab-round-nav').style.display = 'flex';
    document.getElementById('tab-running-status').innerHTML = '<span style="color:var(--success);">Tabulation complete.</span>';
    tabRenderRound(0);
    tabShowSummary();
  }, 50);
}

function tabGoToSeat(seatNum) {
  const idx = tabState.allRounds.findIndex(r => r.seat === seatNum);
  if (idx >= 0) tabRenderRound(idx);
}
function tabPrevRound() { if (tabState.currentRoundIdx > 0) tabRenderRound(tabState.currentRoundIdx - 1); }
function tabNextRound() { if (tabState.currentRoundIdx < tabState.allRounds.length - 1) tabRenderRound(tabState.currentRoundIdx + 1); }
function tabAutoPlay() {
  if (tabState.autoInterval) { clearInterval(tabState.autoInterval); tabState.autoInterval = null; document.getElementById('tab-auto-btn').textContent = 'Auto-Play'; return; }
  document.getElementById('tab-auto-btn').textContent = 'Stop';
  tabState.autoInterval = setInterval(() => {
    if (tabState.currentRoundIdx < tabState.allRounds.length - 1) tabNextRound();
    else { clearInterval(tabState.autoInterval); tabState.autoInterval = null; document.getElementById('tab-auto-btn').textContent = 'Auto-Play'; }
  }, 2000);
}

function tabRenderRound(idx) {
  if (idx < 0 || idx >= tabState.allRounds.length) return;
  tabState.currentRoundIdx = idx;
  const round = tabState.allRounds[idx];
  const cmap = tabState.candidateMap;
  const seatLabel = tabState.seats > 1 ? `Seat ${round.seat} - ` : '';
  document.getElementById('tab-round-label').textContent =
    `${seatLabel}Distribution Round ${round.seatRound} (${idx+1} of ${tabState.allRounds.length})`;
  if (tabState.seats > 1) {
    for (let s = 1; s <= tabState.seats; s++) {
      const btn = document.getElementById('tab-seat-btn-'+s);
      if (btn) btn.className = `btn btn-sm ${s===round.seat ? 'btn-primary' : 'btn-outline'}`;
    }
  }
  const statusDiv = document.getElementById('tab-round-status');
  if (round.elected.length > 0) {
    const name = cmap[round.elected[0]] || round.elected[0];
    statusDiv.textContent = `${name} ELECTED with ${round.voteCounts[round.elected[0]]} votes (${(round.voteCounts[round.elected[0]]/round.totalActive*100).toFixed(1)}%)`;
    statusDiv.style.display = 'block'; statusDiv.style.background = '#e8f5e9'; statusDiv.style.color = '#2e7d32';
  } else if (round.eliminated.length > 0) {
    const names = round.eliminated.map(k => cmap[k]||k).join(', ');
    statusDiv.textContent = `Eliminated: ${names} - ${round.elimReason}`;
    statusDiv.style.display = 'block'; statusDiv.style.background = '#fff3cd'; statusDiv.style.color = '#856404';
  } else if (round.elimReason) {
    statusDiv.textContent = round.elimReason;
    statusDiv.style.display = 'block'; statusDiv.style.background = '#fde8e8'; statusDiv.style.color = '#c0392b';
  } else { statusDiv.style.display = 'none'; }

  const chartDiv = document.getElementById('tab-chart');
  const sorted = round.active.slice().sort((a,b) => (round.voteCounts[b]||0) - (round.voteCounts[a]||0));
  const maxVotes = Math.max(1, ...Object.values(round.voteCounts));
  chartDiv.innerHTML = sorted.map(k => {
    const votes = round.voteCounts[k] || 0;
    const pct = round.totalActive ? (votes/round.totalActive*100).toFixed(1) : '0';
    const isElected = round.elected.includes(k);
    const isElim = round.eliminated.includes(k);
    const color = isElected ? 'var(--success)' : isElim ? 'var(--danger)' : 'var(--navy)';
    return `<div style="display:flex;align-items:center;gap:.75rem;margin-bottom:.4rem;opacity:${isElim?'.5':'1'};">
      <span style="width:160px;font-size:.85rem;font-weight:600;text-align:right;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${k}. ${cmap[k]||k}</span>
      <div style="flex:1;height:28px;background:#f0f0f0;border-radius:4px;overflow:hidden;">
        <div style="height:100%;width:${(votes/maxVotes*100).toFixed(1)}%;background:${color};border-radius:4px;transition:width .4s;"></div>
      </div>
      <span style="width:80px;font-size:.85rem;text-align:right;">${votes} (${pct}%)</span>
      ${isElected ? '<span class="badge badge-success">ELECTED</span>' : ''}
      ${isElim ? '<span class="badge" style="background:#fde8e8;color:#c0392b;">OUT</span>' : ''}
    </div>`;
  }).join('');
  if (round.majorityNeeded) {
    chartDiv.innerHTML += `<div style="font-size:.8rem;color:var(--text-muted);margin-top:.5rem;">Majority needed: ${round.majorityNeeded} of ${round.totalActive} active votes &middot; ${round.exhausted} exhausted</div>`;
  }

  document.getElementById('tab-table-wrap').innerHTML = `<table><thead><tr><th>Candidate</th><th>Votes</th><th>%</th><th>Status</th></tr></thead><tbody>${
    sorted.map(k => {
      const votes = round.voteCounts[k] || 0;
      const pct = round.totalActive ? (votes/round.totalActive*100).toFixed(1) : '0';
      const isE = round.elected.includes(k); const isX = round.eliminated.includes(k);
      return `<tr style="${isX?'opacity:.5;':''}"><td><strong>${k}.</strong> ${cmap[k]||k}</td><td>${votes}</td><td>${pct}%</td><td>${isE ? '<span class="badge badge-success">ELECTED</span>' : isX ? '<span class="badge" style="background:#fde8e8;color:#c0392b;">ELIMINATED</span>' : 'Active'}</td></tr>`;
    }).join('')
  }</tbody></table>`;
}

function tabShowSummary() {
  document.getElementById('tab-summary-card').style.display = 'block';
  const cmap = tabState.candidateMap;
  let html = '<div style="margin-bottom:1rem;">';
  if (tabState.electedAll.length > 0) {
    html += `<h4 style="color:var(--success);margin-bottom:.5rem;">Elected (${tabState.electedAll.length} of ${tabState.seats} seat${tabState.seats!==1?'s':''})</h4><ol>`;
    tabState.electedAll.forEach((k, i) => {
      const round = tabState.allRounds.find(r => r.elected.includes(k));
      html += `<li style="margin-bottom:.3rem;"><strong>${cmap[k] || k}</strong> - Seat ${round?.seat || i+1}, Round ${round?.seatRound || '?'} (${round?.voteCounts[k] || '?'} votes, ${round?.totalActive ? ((round.voteCounts[k]/round.totalActive*100).toFixed(1))+'%' : '?'})</li>`;
    });
    html += '</ol>';
  } else { html += '<p style="color:var(--danger);">No candidates elected.</p>'; }
  html += `</div><div style="margin-top:1rem;padding-top:.75rem;border-top:1px solid var(--border);">
    <h4 style="margin-bottom:.5rem;">Credentials Report</h4>
    <p style="font-size:.85rem;">Total ballots: ${tabState.ballots.length + tabState.spoiled.length} &middot; Valid: ${tabState.ballots.length} &middot; Flagged: ${tabState.spoiled.length} &middot; Surrendered on file: ${STATE.surrendered.length}</p>
  </div>`;
  html += `<div style="margin-top:.75rem;font-size:.85rem;color:var(--text-muted);">Total distribution rounds: ${tabState.allRounds.length}</div>`;
  document.getElementById('tab-summary-content').innerHTML = html;
}

</script>
</body>
</html>"""



# ═════════════════════════════════════════════════════════════════════════
#  HTTP SERVER + API HANDLER
# ═════════════════════════════════════════════════════════════════════════

class AdminHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        pass

    def send_json(self, data, code=200):
        body = json.dumps(data).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_body(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        return json.loads(raw) if raw else {}

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/" or path == "":
            body = HTML_TEMPLATE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path == "/api/status":
            elections = get_elections()
            self.send_json({
                "status": "ok",
                "elections": elections,
                "nominees": get_nominees(),
                "delegates": get_delegates(),
                "surrendered": get_surrendered(),
            })
            return
        if path.startswith("/qr/"):
            filename = path[4:]
            qr_path = QR_DIR / filename
            if qr_path.exists():
                with open(qr_path, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            else:
                self.send_json({"status": "error", "message": "Not found"}, 404)
            return
        self.send_json({"status": "error", "message": "Not found"}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            body = self.read_body()
        except Exception:
            body = {}

        if path == "/api/elections/save":
            elections = get_elections()
            key = body.get("key", "")
            if not key:
                self.send_json({"status": "error", "message": "Key required"})
                return
            elections[key] = body
            save_elections(elections)
            self.send_json({"status": "ok"})
            return

        if path == "/api/elections/delete":
            elections = get_elections()
            key = body.get("key", "")
            elections.pop(key, None)
            save_elections(elections)
            noms = get_nominees()
            noms.pop(key, None)
            save_json(NOMINEES_FILE, noms)
            self.send_json({"status": "ok"})
            return

        if path == "/api/nominees/add":
            noms = get_nominees()
            key = body.get("key", "")
            if key not in noms:
                noms[key] = []
            noms[key].append({"letter": body["letter"], "name": body["name"]})
            noms[key].sort(key=lambda n: n["letter"])
            save_json(NOMINEES_FILE, noms)
            self.send_json({"status": "ok"})
            return

        if path == "/api/nominees/remove":
            noms = get_nominees()
            key = body.get("key", "")
            letter = body.get("letter", "")
            if key in noms:
                noms[key] = [n for n in noms[key] if n["letter"] != letter]
                save_json(NOMINEES_FILE, noms)
            self.send_json({"status": "ok"})
            return

        if path == "/api/delegates/add":
            dels = get_delegates()
            dels.append({
                "id": body["id"], "last": body["last"], "first": body["first"],
                "email": body.get("email", ""), "phone": body.get("phone", ""),
            })
            save_json(DELEGATES_FILE, dels)
            self.send_json({"status": "ok"})
            return

        if path == "/api/delegates/import":
            rows = body.get("delegates", [])
            save_json(DELEGATES_FILE, rows)
            self.send_json({"status": "ok", "delegates": rows})
            return

        if path == "/api/surrendered/add":
            surr = get_surrendered()
            surr.append({
                "id": body["id"], "status": body.get("status", "surrendered"),
                "time": body.get("time", ""), "note": body.get("note", ""),
            })
            save_json(SURRENDERED_FILE, surr)
            self.send_json({"status": "ok"})
            return

        if path == "/api/surrendered/remove":
            surr = get_surrendered()
            surr = [s for s in surr if s["id"] != body.get("id", "")]
            save_json(SURRENDERED_FILE, surr)
            self.send_json({"status": "ok"})
            return

        if path == "/api/generate":
            key = body.get("key", "")
            duration = body.get("duration")
            elections = get_elections()
            if key not in elections:
                self.send_json({"status": "error", "message": "Unknown election"})
                return
            noms = get_nominees().get(key, [])
            if not noms:
                self.send_json({"status": "error", "message": "No nominees"})
                return
            election = elections[key]
            ballot_url = build_ballot_url(key, noms, duration)
            admin_url = build_ballot_url(key, noms, duration, admin=True)
            qr_b64 = generate_qr_base64(ballot_url)
            qr_file = save_qr_png(ballot_url, f"{key}_ballot_qr.png")
            display_path = DISPLAY_DIR / f"{key}_display.html"
            write_projection_html(str(display_path), election, noms, ballot_url, timer_minutes=duration)
            self.send_json({
                "status": "ok", "ballot_url": ballot_url, "admin_url": admin_url,
                "qr": qr_b64, "qr_file": str(qr_file) if qr_file else None,
                "display_file": str(display_path),
            })
            return

        if path == "/api/generate-practice":
            practice_url = f"{PRACTICE_BASE_URL}?mode=practice"
            qr_b64 = generate_qr_base64(practice_url)
            save_qr_png(practice_url, "practice_ballot_qr.png")
            display_path = DISPLAY_DIR / "practice_display.html"
            write_practice_display_html(str(display_path), practice_url)
            self.send_json({"status": "ok", "ballot_url": practice_url, "qr": qr_b64})
            return

        if path == "/api/generate-assisted":
            key = body.get("key", "")
            elections = get_elections()
            if key not in elections:
                self.send_json({"status": "error", "message": "Unknown election"})
                return
            noms = get_nominees().get(key, [])
            if not noms:
                self.send_json({"status": "error", "message": "No nominees"})
                return
            url = build_assisted_url(key, noms)
            self.send_json({"status": "ok", "url": url})
            return

        if path == "/api/voting/open":
            key = body.get("key", "")
            env = body.get("env", "test")
            floor_count = body.get("floorCount", 0)
            payload = {"action": "open", "election": key, "floorCount": floor_count}
            result = apps_script_post(payload, use_test=(env == "test"))
            self.send_json({"status": "ok", "data": result})
            return

        if path == "/api/voting/close":
            key = body.get("key", "")
            env = body.get("env", "test")
            payload = {"action": "close", "election": key}
            result = apps_script_post(payload, use_test=(env == "test"))
            self.send_json({"status": "ok", "data": result})
            return

        if path == "/api/voting/status":
            key = body.get("key", "")
            env = body.get("env", "test")
            result = apps_script_get(key, use_test=(env == "test"))
            self.send_json({"status": "ok", "data": result})
            return

        if path == "/api/push-setup":
            key = body.get("key", "")
            env = body.get("env", "test")
            elections = get_elections()
            if key not in elections:
                self.send_json({"status": "error", "message": "Unknown election"})
                return
            election = elections[key]
            noms = get_nominees().get(key, [])
            duration = body.get("duration")
            ballot_type = body.get("ballotType", election.get("type", "ranked"))
            floor_count = body.get("floorCount", 0)
            ballot_url = build_ballot_url(key, noms, duration) if noms else ""
            admin_url = build_ballot_url(key, noms, duration, admin=True) if noms else ""
            assisted_url = build_assisted_url(key, noms, duration) if noms else ""
            payload = {
                "action": "setupElection", "election": key, "label": election["label"],
                "seats": election["seats"], "ballotType": ballot_type,
                "candidates": ", ".join(f"{n['letter']}. {n['name']}" for n in noms),
                "ballotUrl": ballot_url, "adminUrl": admin_url, "assistedUrl": assisted_url,
                "duration": duration, "floorCount": floor_count,
                "inactiveCredentials": ", ".join(s["id"] for s in get_surrendered()),
            }
            result = apps_script_post(payload, use_test=(env == "test"))
            self.send_json({"status": "ok", "data": result})
            return

        if path == "/api/mark-live":
            key = body.get("key", "")
            env = body.get("env", "test")
            payload = {
                "action": "updateSetup", "election": key,
                "field": "Candidates Live At",
                "value": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            result = apps_script_post(payload, use_test=(env == "test"))
            self.send_json({"status": "ok", "data": result})
            return

        if path == "/api/export-ballots":
            key = body.get("key", "")
            env = body.get("env", "test")
            if not key:
                self.send_json({"status": "error", "message": "Election key required"})
                return
            payload = {"action": "exportBallots", "electionKey": key}
            result = apps_script_post(payload, use_test=(env == "test"))
            if isinstance(result, dict) and result.get("status") == "ok":
                self.send_json({
                    "status": "ok",
                    "headers": result.get("headers", []),
                    "rows": result.get("rows", []),
                })
            else:
                msg = result.get("message", "Failed to export") if isinstance(result, dict) else str(result)
                self.send_json({"status": "error", "message": msg})
            return

        self.send_json({"status": "error", "message": "Unknown route"}, 404)


def main():
    print(f"\n{'='*60}")
    print(f"  3rd CD Convention Admin Dashboard")
    print(f"  http://localhost:{PORT}")
    print(f"{'='*60}")
    print(f"  Data folder : {DATA_DIR}")
    print(f"  QR folder   : {QR_DIR}")
    print(f"  Display folder: {DISPLAY_DIR}")
    print(f"  Elections   : {len(get_elections())}")
    print(f"  Nominees    : {sum(len(v) for v in get_nominees().values())}")
    print(f"  Delegates   : {len(get_delegates())}")
    print(f"  Surrendered : {len(get_surrendered())}")
    print(f"{'='*60}")
    print(f"  Press Ctrl+C to stop\n")

    server = HTTPServer(("127.0.0.1", PORT), AdminHandler)

    def open_browser():
        import time; time.sleep(0.5)
        webbrowser.open(f"http://localhost:{PORT}")
    threading.Thread(target=open_browser, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Dashboard stopped.")
        server.server_close()


if __name__ == "__main__":
    main()
