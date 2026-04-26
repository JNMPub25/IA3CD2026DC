#!/usr/bin/env python3
"""
District 3 Convention — Day-Of Setup Script
Manages nominees, delegate roster, surrendered credentials,
generates QR codes, and produces projection display files.

INSTALL REQUIREMENTS (run once on your machine):
    pip install qrcode[pil]

Run this script from the project folder:
    python district3_setup.py
"""

import os
import json
import csv
import sys
import re
import urllib.parse
from datetime import datetime, timedelta
from pathlib import Path

# ── Try importing qrcode (graceful fallback if not installed) ──────────────
try:
    import qrcode
    from qrcode.image.pure import PyPNGImage
    QRCODE_AVAILABLE = True
except ImportError:
    QRCODE_AVAILABLE = False

# ── Folder layout ─────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
DATA_DIR   = BASE_DIR / "setup_data"
QR_DIR     = BASE_DIR / "qr_codes"
DISPLAY_DIR = BASE_DIR / "projection_displays"

for d in [DATA_DIR, QR_DIR, DISPLAY_DIR]:
    d.mkdir(exist_ok=True)

# ── Data-file helpers ──────────────────────────────────────────────────────
NOMINEES_FILE    = DATA_DIR / "nominees.json"
DELEGATES_FILE   = DATA_DIR / "delegates.json"
SURRENDERED_FILE = DATA_DIR / "surrendered_delegates.json"

def load_json(path, default):
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return default

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

# ── Election definitions ───────────────────────────────────────────────────
ELECTIONS = {
    "1": {"key": "scc-w",  "label": "SCC — Women",              "seats": 4,  "type": "ranked"},
    "2": {"key": "scc-m",  "label": "SCC — Men / Non-Binary",   "seats": 4,  "type": "ranked"},
    "3": {"key": "dei",    "label": "DEI Committee Chair",       "seats": 1,  "type": "ranked"},
    "4": {"key": "scc-com","label": "State Convention Committee","seats": 14, "type": "slate"},
}

# ── QR / URL base ─────────────────────────────────────────────────────────
# Replace with your actual GitHub Pages URL once the ballot is deployed.
BALLOT_BASE_URL   = "https://jnmpub25.github.io/IA3CD2026DC/ballot.html"
PRACTICE_BASE_URL = "https://jnmpub25.github.io/IA3CD2026DC/practice_ballot.html"

# Fictional candidates shown in the practice ballot (must match practice_ballot.html)
PRACTICE_CANDIDATES = [
    {"letter": "A", "name": "Adams, Carol"},
    {"letter": "B", "name": "Bennett, Diane"},
    {"letter": "C", "name": "Chen, Megan"},
    {"letter": "D", "name": "Davis, Rachel"},
    {"letter": "E", "name": "Evans, Patricia"},
    {"letter": "F", "name": "Flores, Sandra"},
]

# ─────────────────────────────────────────────────────────────────────────
#  SECTION 1 — NOMINEE MANAGEMENT
# ─────────────────────────────────────────────────────────────────────────

def manage_nominees():
    nominees = load_json(NOMINEES_FILE, {e["key"]: [] for e in ELECTIONS.values()})

    while True:
        print("\n╔══════════════════════════════════════╗")
        print("║       NOMINEE MANAGEMENT             ║")
        print("╚══════════════════════════════════════╝")
        for num, e in ELECTIONS.items():
            count = len(nominees.get(e["key"], []))
            print(f"  {num}. {e['label']}  ({count} nominees)")
        print("  5. Back to Main Menu")
        choice = input("\nSelect election to manage: ").strip()

        if choice == "5":
            break
        if choice not in ELECTIONS:
            print("  ⚠  Invalid choice.")
            continue

        election = ELECTIONS[choice]
        key = election["key"]
        if key not in nominees:
            nominees[key] = []

        _edit_nominees(nominees, election)
        save_json(NOMINEES_FILE, nominees)


def _edit_nominees(nominees, election):
    key = election["key"]
    while True:
        print(f"\n── {election['label']} ──")
        current = nominees[key]
        if current:
            for n in current:
                print(f"    {n['letter']}. {n['name']}")
        else:
            print("    (no nominees yet)")
        print()
        print("  A  Add nominee")
        print("  R  Remove nominee")
        print("  B  Back")
        action = input("Action: ").strip().upper()

        if action == "B":
            break
        elif action == "A":
            _add_nominee(nominees, key)
        elif action == "R":
            _remove_nominee(nominees, key)
        else:
            print("  ⚠  Type A, R, or B.")


def _add_nominee(nominees, key):
    print("\nEnter candidate letter (A, B, C …) and name.")
    print("Tip: letters are assigned in advance; names appear on the ballot as entered.")
    letter = input("  Candidate letter: ").strip().upper()
    if not letter.isalpha() or len(letter) != 1:
        print("  ⚠  Letter must be a single alphabetic character (A, B, C …).")
        return
    name = input("  Candidate name: ").strip()
    if not name:
        print("  ⚠  Name cannot be blank.")
        return
    # Prevent duplicate letters
    if any(n["letter"] == letter for n in nominees[key]):
        print(f"  ⚠  Candidate {letter} already exists.")
        return
    nominees[key].append({"letter": letter, "name": name})
    nominees[key].sort(key=lambda x: x["letter"])
    print(f"  ✓  Added: {letter}. {name}")


def _remove_nominee(nominees, key):
    letter = input("  Enter candidate letter to remove: ").strip().upper()
    before = len(nominees[key])
    nominees[key] = [n for n in nominees[key] if n["letter"] != letter]
    if len(nominees[key]) < before:
        print(f"  ✓  Candidate {letter} removed.")
    else:
        print(f"  ⚠  Candidate {letter} not found.")


# ─────────────────────────────────────────────────────────────────────────
#  SECTION 2 — DELEGATE ROSTER
# ─────────────────────────────────────────────────────────────────────────

def manage_delegates():
    delegates = load_json(DELEGATES_FILE, [])

    while True:
        print("\n╔══════════════════════════════════════╗")
        print("║       DELEGATE ROSTER                ║")
        print("╚══════════════════════════════════════╝")
        print(f"  Delegates loaded: {len(delegates)}")
        print()
        print("  1. Import delegate list from CSV file")
        print("  2. View delegate list")
        print("  3. Add single delegate")
        print("  4. Back to Main Menu")
        choice = input("\nChoice: ").strip()

        if choice == "4":
            break
        elif choice == "1":
            delegates = _import_delegates_csv(delegates)
            save_json(DELEGATES_FILE, delegates)
        elif choice == "2":
            _view_delegates(delegates)
        elif choice == "3":
            delegates = _add_single_delegate(delegates)
            save_json(DELEGATES_FILE, delegates)
        else:
            print("  ⚠  Invalid choice.")


def _pick_column(headers, prompt, required=True):
    """
    Display the CSV headers as a numbered list and let the user
    pick by number.  Returns the original header string, or None
    if the user skips (only allowed when required=False).
    """
    print(f"\n  {prompt}")
    for i, h in enumerate(headers, 1):
        print(f"    {i:>2}. {h}")
    if not required:
        print(f"    {0:>2}. Skip (not available in this file)")
    while True:
        raw = input("     Enter number: ").strip()
        if raw == "0" and not required:
            return None
        if raw.isdigit() and 1 <= int(raw) <= len(headers):
            return headers[int(raw) - 1]
        print(f"  ⚠  Enter a number between {'0 and' if not required else '1 and'} {len(headers)}.")


def _import_delegates_csv(existing):
    """
    Import delegate roster from any CSV file.
    The user identifies which column maps to each field by
    choosing from a numbered list of the headers found in the file.
    Only the delegate ID column is required; all others may be skipped.
    """
    print("\nDelegate CSV Import")
    print("──────────────────────────────────────────────────────")
    print("You will be shown the column headers found in your file")
    print("and asked to identify which column contains each field.")
    print()
    path_str = input("Enter full path to CSV file (or drag-and-drop here): ").strip().strip('"')
    path = Path(path_str)

    if not path.exists():
        print(f"  ⚠  File not found: {path}")
        return existing

    # ── Read headers from file ────────────────────────────────────────────
    # Use cp1252 (Windows-1252) — handles ASCII, BOM-less UTF-8 printable
    # chars, and Windows special characters (en-dashes, curly quotes, etc.)
    # that appear in phone/name fields exported from Excel on Windows.
    with open(path, "r", encoding="cp1252", newline="") as f:
        reader = csv.DictReader(f)
        headers = list(reader.fieldnames or [])

    if not headers:
        print("  ⚠  No headers found in file. Is this a valid CSV?")
        return existing

    print(f"\n  Found {len(headers)} column(s) in your file:")
    for i, h in enumerate(headers, 1):
        print(f"    {i:>2}. {h}")

    # ── Map columns ───────────────────────────────────────────────────────
    col_id = _pick_column(headers,
                          "Enter the number for the column that contains the delegate ID / number:",
                          required=True)

    combined = input("\n  Are the first and last name in the same column? (Y/N): ").strip().upper()
    if combined == "Y":
        col_name  = _pick_column(headers,
                                 "Enter the number for the column that contains the delegate's name:",
                                 required=False)
        col_first = None
        col_last  = None
    else:
        col_name  = None
        col_last  = _pick_column(headers,
                                 "Enter the number for the column that contains the delegate's Last Name:",
                                 required=False)
        col_first = _pick_column(headers,
                                 "Enter the number for the column that contains the delegate's First Name:",
                                 required=False)

    col_phone = _pick_column(headers,
                             "Enter the number for the column that contains the delegate's phone number:",
                             required=False)
    col_email = _pick_column(headers,
                             "Enter the number for the column that contains the delegate's email address:",
                             required=False)

    print()

    # ── Read and map rows ─────────────────────────────────────────────────
    imported = []
    skipped  = 0
    with open(path, "r", encoding="cp1252", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            num = row.get(col_id, "").strip()
            if not num:
                skipped += 1
                continue

            # Resolve name fields
            if col_name:
                # Single combined column — store as-is; split on last space for first/last
                full  = row.get(col_name, "").strip()
                parts = full.rsplit(" ", 1)
                first = parts[0] if len(parts) > 1 else full
                last  = parts[1] if len(parts) > 1 else ""
            else:
                last  = row.get(col_last,  "").strip() if col_last  else ""
                first = row.get(col_first, "").strip() if col_first else ""

            # Display name stored as "Last, First" for roster and reports
            if last and first:
                display_name = f"{last}, {first}"
            elif last:
                display_name = last
            else:
                display_name = first

            delegate = {
                "delegate_number": num,
                "first_name":      first,
                "last_name":       last,
                "display_name":    display_name,
                "phone": row.get(col_phone, "").strip() if col_phone else "",
                "email": row.get(col_email, "").strip() if col_email else "",
            }
            imported.append(delegate)

    # Merge: imported replaces existing entries with same delegate_number
    existing_by_num = {d["delegate_number"]: d for d in existing}
    for d in imported:
        existing_by_num[d["delegate_number"]] = d
    merged = sorted(existing_by_num.values(), key=lambda x: x["delegate_number"])

    print(f"  ✓  Imported {len(imported)} delegates ({skipped} rows skipped).")
    print(f"     Total roster: {len(merged)}")
    return merged


def _view_delegates(delegates):
    if not delegates:
        print("  (no delegates loaded)")
        return
    print(f"\n{'#':<8} {'Name (Last, First)':<35} {'Email':<30} {'Phone'}")
    print("─" * 90)
    for d in delegates[:50]:
        name = d.get("display_name") or f"{d.get('last_name','')} {d.get('first_name','')}".strip()
        print(f"{d['delegate_number']:<8} {name:<35} {d.get('email',''):<30} {d.get('phone','')}")
    if len(delegates) > 50:
        print(f"  … and {len(delegates) - 50} more. Full list saved to setup_data/delegates.json")


def _add_single_delegate(delegates):
    print("\nAdd Single Delegate")
    num    = input("  Delegate ID (e.g. D-6001 or A-6063): ").strip()
    last   = input("  Last name       : ").strip()
    first  = input("  First name      : ").strip()
    email  = input("  Email           : ").strip()
    phone  = input("  Phone           : ").strip()
    if not num:
        print("  ⚠  Delegate ID required.")
        return delegates
    display_name = f"{last}, {first}" if last and first else (last or first)
    # Replace if exists
    delegates = [d for d in delegates if d["delegate_number"] != num]
    delegates.append({"delegate_number": num, "last_name": last,
                       "first_name": first, "display_name": display_name,
                       "email": email, "phone": phone})
    delegates.sort(key=lambda x: x["delegate_number"])
    print(f"  ✓  Delegate {num} ({display_name}) added.")
    return delegates


# ─────────────────────────────────────────────────────────────────────────
#  SECTION 3 — SURRENDERED CREDENTIALS
# ─────────────────────────────────────────────────────────────────────────

def manage_surrendered():
    surrendered = load_json(SURRENDERED_FILE, [])

    while True:
        print("\n╔══════════════════════════════════════╗")
        print("║     SURRENDERED CREDENTIALS          ║")
        print("╚══════════════════════════════════════╝")
        if surrendered:
            print(f"  {'Delegate':<12} {'Surrender Time':<22} {'Name'}")
            print("  " + "─" * 55)
            for s in surrendered:
                print(f"  {s['delegate_number']:<12} {s['surrender_time']:<22} {s.get('name','')}")
        else:
            print("  (no surrendered delegates)")
        print()
        print("  1. Record surrendered delegate")
        print("  2. Remove entry (delegate reinstated)")
        print("  3. Back to Main Menu")
        choice = input("\nChoice: ").strip()

        if choice == "3":
            break
        elif choice == "1":
            surrendered = _record_surrender(surrendered)
            save_json(SURRENDERED_FILE, surrendered)
        elif choice == "2":
            surrendered = _remove_surrender(surrendered)
            save_json(SURRENDERED_FILE, surrendered)
        else:
            print("  ⚠  Invalid choice.")


def _record_surrender(surrendered):
    delegates = load_json(DELEGATES_FILE, [])
    by_num    = {d["delegate_number"]: d for d in delegates}

    num = input("  Delegate ID (e.g. D-6001 or A-6063): ").strip()
    # Look up name if we have it
    delegate_info = by_num.get(num, {})
    name = f"{delegate_info.get('first_name','')} {delegate_info.get('last_name','')}".strip()

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    custom  = input(f"  Surrender time [{now_str}] (press Enter for now): ").strip()
    surrender_time = custom if custom else now_str

    # Remove any existing entry for this delegate first
    surrendered = [s for s in surrendered if s["delegate_number"] != num]
    surrendered.append({
        "delegate_number": num,
        "name":            name,
        "surrender_time":  surrender_time,
        "recorded_by_script": datetime.now().isoformat()
    })
    surrendered.sort(key=lambda x: x["delegate_number"])
    print(f"  ✓  Delegate {num} ({name}) marked as surrendered at {surrender_time}.")
    print("     Any ballot submitted after this time will be flagged as")
    print("     'Spoiled — Surrendered Credentials' in the exceptions report.")
    return surrendered


def _remove_surrender(surrendered):
    num = input("  Delegate ID to reinstate (e.g. D-6001 or A-6063): ").strip()
    before = len(surrendered)
    surrendered = [s for s in surrendered if s["delegate_number"] != num]
    if len(surrendered) < before:
        print(f"  ✓  Delegate {num} reinstated — their ballot will no longer be flagged.")
    else:
        print(f"  ⚠  Delegate {num} not in surrendered list.")
    return surrendered


# ─────────────────────────────────────────────────────────────────────────
#  SECTION 4 — GENERATE QR CODE & PROJECTION DISPLAY
# ─────────────────────────────────────────────────────────────────────────

def generate_qr_and_display():
    nominees = load_json(NOMINEES_FILE, {e["key"]: [] for e in ELECTIONS.values()})

    print("\n╔══════════════════════════════════════╗")
    print("║   GENERATE QR CODE & DISPLAY SLIDE  ║")
    print("╚══════════════════════════════════════╝")
    for num, e in ELECTIONS.items():
        count = len(nominees.get(e["key"], []))
        print(f"  {num}. {e['label']}  ({count} nominees)")
    print("  5. Back to Main Menu")
    choice = input("\nSelect election: ").strip()

    if choice == "5":
        return
    if choice not in ELECTIONS:
        print("  ⚠  Invalid choice.")
        return

    election = ELECTIONS[choice]
    key      = election["key"]
    noms     = nominees.get(key, [])

    if not noms:
        print(f"  ⚠  No nominees for {election['label']}. Please add nominees first.")
        return

    # Build URL
    candidate_param = "-".join(
        f"{n['letter']}-{urllib.parse.quote(n['name'], safe='')}"
        for n in noms
    )
    url = (f"{BALLOT_BASE_URL}"
           f"?election={key}"
           f"&candidates={candidate_param}"
           f"&seats={election['seats']}")

    print(f"\n  Ballot URL:\n  {url}\n")

    # ── Generate QR code PNG ──────────────────────────────────────────────
    qr_path = QR_DIR / f"qr_{key}.png"
    if QRCODE_AVAILABLE:
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=4,
        )
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        img.save(str(qr_path))
        print(f"  ✓  QR code saved: {qr_path}")
    else:
        print("  ⚠  qrcode library not installed — QR PNG not generated.")
        print("     Install with:  pip install qrcode[pil]")
        print("     The projection display will still generate a QR code via JavaScript.")

    # ── Generate projection display HTML ─────────────────────────────────
    display_path = DISPLAY_DIR / f"display_{key}.html"
    _write_projection_html(display_path, election, noms, url)
    print(f"  ✓  Phase 1 projection display saved: {display_path}")

    # ── Phase 2 (timed) ballot URL ───────────────────────────────────────
    p2 = input("\n  Generate a Phase 2 (timed) ballot URL? (Y/N): ").strip().upper()
    if p2 == "Y":
        while True:
            try:
                minutes = int(input("  Duration in minutes for Phase 2: ").strip())
                if minutes > 0:
                    break
                print("  ⚠  Enter a positive number of minutes.")
            except ValueError:
                print("  ⚠  Enter a whole number of minutes (e.g. 10).")

        deadline_dt  = datetime.now() + timedelta(minutes=minutes)
        deadline_iso = urllib.parse.quote(deadline_dt.isoformat(), safe="")
        url_phase2   = f"{url}&deadline={deadline_iso}"

        # QR code for Phase 2
        qr_path2 = QR_DIR / f"qr_{key}_phase2.png"
        if QRCODE_AVAILABLE:
            qr2 = qrcode.QRCode(
                version=None,
                error_correction=qrcode.constants.ERROR_CORRECT_M,
                box_size=10,
                border=4,
            )
            qr2.add_data(url_phase2)
            qr2.make(fit=True)
            img2 = qr2.make_image(fill_color="black", back_color="white")
            img2.save(str(qr_path2))
            print(f"  ✓  Phase 2 QR code saved: {qr_path2}")
        else:
            print("  ⚠  qrcode library not available — Phase 2 QR PNG not generated.")
            print("     The projection display will generate the QR code via JavaScript.")

        # Projection display with Phase 2 badge
        display_path2 = DISPLAY_DIR / f"display_{key}_phase2.html"
        _write_projection_html(display_path2, election, noms, url_phase2,
                               phase2_minutes=minutes)
        print(f"  ✓  Phase 2 projection display saved: {display_path2}")

        print()
        print(f"  Phase 1 URL : {url}")
        print(f"  Phase 2 URL : {url_phase2}")
        print(f"  Deadline    : {deadline_dt.strftime('%H:%M:%S  %B %d, %Y')}")
    else:
        print(f"\n  Ballot URL (Phase 1 only): {url}")

    print("\n  To project: open the HTML file in Chrome, then press F11 for fullscreen.")


def _write_projection_html(path, election, nominees, url, phase2_minutes=None):
    """
    Generate a fullscreen HTML projection display showing:
    - Election title and seat count (plus PHASE 2 badge when phase2_minutes is set)
    - Lettered candidate list
    - QR code (generated in-browser via qrcodejs — works offline)
    - Ballot URL as text
    """
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
    background: #003366;
    color: #ffffff;
    font-family: 'Segoe UI', Arial, sans-serif;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 100vh;
    padding: 2rem;
  }}
  .header {{
    text-align: center;
    margin-bottom: 2rem;
  }}
  .header h1 {{
    font-size: 3rem;
    font-weight: 700;
    letter-spacing: 0.04em;
    text-transform: uppercase;
  }}
  .header p {{
    font-size: 1.4rem;
    opacity: 0.85;
    margin-top: 0.4rem;
  }}
  .content {{
    display: flex;
    gap: 4rem;
    align-items: flex-start;
    justify-content: center;
    flex-wrap: wrap;
    width: 100%;
    max-width: 1200px;
  }}
  .candidates {{
    flex: 1;
    min-width: 300px;
  }}
  .candidates h2 {{
    font-size: 1.4rem;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    opacity: 0.7;
    margin-bottom: 1rem;
    border-bottom: 2px solid rgba(255,255,255,0.3);
    padding-bottom: 0.4rem;
  }}
  .candidates ol {{
    list-style: none;
    padding: 0;
  }}
  .candidates li {{
    font-size: 1.8rem;
    padding: 0.5rem 0;
    border-bottom: 1px solid rgba(255,255,255,0.15);
    display: flex;
    align-items: center;
    gap: 1rem;
  }}
  .num {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    background: rgba(255,255,255,0.2);
    border-radius: 50%;
    width: 2.2rem;
    height: 2.2rem;
    font-size: 1.1rem;
    font-weight: 700;
    flex-shrink: 0;
  }}
  .qr-panel {{
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 1rem;
  }}
  .qr-panel h2 {{
    font-size: 1.4rem;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    opacity: 0.7;
  }}
  #qrcode canvas, #qrcode img {{
    border: 8px solid white;
    border-radius: 8px;
  }}
  .url-box {{
    background: rgba(255,255,255,0.15);
    border-radius: 6px;
    padding: 0.8rem 1.2rem;
    font-size: 0.85rem;
    word-break: break-all;
    max-width: 340px;
    text-align: center;
    opacity: 0.8;
  }}
  .seats-badge {{
    display: inline-block;
    background: #FFD700;
    color: #003366;
    font-weight: 700;
    font-size: 1rem;
    padding: 0.3rem 1rem;
    border-radius: 20px;
    margin-top: 0.5rem;
  }}
  .phase2-badge {{
    display: inline-block;
    background: #c0392b;
    color: #fff;
    font-weight: 700;
    font-size: 1.1rem;
    padding: 0.35rem 1.2rem;
    border-radius: 20px;
    margin-top: 0.6rem;
    letter-spacing: 0.06em;
  }}
</style>
</head>
<body>
<div class="header">
  <h1>3rd Congressional District</h1>
  <p>Iowa Democratic Party — District Convention {datetime.now().year}</p>
  <div class="seats-badge">{election['label'].upper()} &nbsp;|&nbsp; {election['seats']} Seat{'s' if election['seats'] != 1 else ''}</div>
  {f'<div class="phase2-badge">⏱ PHASE 2 — {phase2_minutes} MINUTES</div>' if phase2_minutes else ''}
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

<!-- qrcodejs — pure JavaScript QR generator, works offline after first load -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"
        integrity="sha512-CNgIRecGo7nphbeZ04Sc13ka07paqdeTu0WR1IM4kNcpmBAXAIn1KFj2XSQO0n3AI4wgv77RX9Qc6YGcOiMaA=="
        crossorigin="anonymous" referrerpolicy="no-referrer"></script>
<script>
  new QRCode(document.getElementById("qrcode"), {{
    text: "{url}",
    width: 300,
    height: 300,
    colorDark: "#000000",
    colorLight: "#ffffff",
    correctLevel: QRCode.CorrectLevel.M
  }});
</script>
</body>
</html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)


# ─────────────────────────────────────────────────────────────────────────
#  SECTION 5 — PRACTICE BALLOT QR CODE & DISPLAY SLIDE
# ─────────────────────────────────────────────────────────────────────────

def generate_practice_qr():
    """
    Generate a QR code and projection display for the delegate practice ballot.
    The practice ballot uses fictional candidates hardcoded in practice_ballot.html.
    No election parameters are needed in the URL.
    """
    print("\n╔══════════════════════════════════════╗")
    print("║   PRACTICE BALLOT QR & DISPLAY SLIDE ║")
    print("╚══════════════════════════════════════╝")
    print()
    print("  This generates a QR code for the delegate practice ballot.")
    print(f"  Fictional candidates used: {', '.join(c['name'] for c in PRACTICE_CANDIDATES)}")
    print()

    url = PRACTICE_BASE_URL
    print(f"  Practice ballot URL:\n  {url}\n")

    # ── Generate QR code PNG ──────────────────────────────────────────────
    qr_path = QR_DIR / "qr_practice.png"
    if QRCODE_AVAILABLE:
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=4,
        )
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        img.save(str(qr_path))
        print(f"  ✓  QR code saved: {qr_path}")
    else:
        print("  ⚠  qrcode library not installed — QR PNG not generated.")
        print("     Install with:  pip install qrcode[pil]")
        print("     The projection display will still show the QR code via JavaScript.")

    # ── Generate projection display HTML ─────────────────────────────────
    display_path = DISPLAY_DIR / "display_practice.html"
    _write_practice_display_html(display_path, url)
    print(f"  ✓  Projection display saved: {display_path}")
    print()
    print("  To project: open the HTML file in Chrome, then press F11 for fullscreen.")
    print("  Delegates scan the QR code or use the URL to open the practice ballot.")


def _write_practice_display_html(path, url):
    """
    Fullscreen projection display for the practice session.
    Shows the fictional candidate list alongside the QR code.
    Styled in gold/amber to distinguish it from the real election displays.
    """
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
  body {{
    background: #3e2700;
    color: #ffffff;
    font-family: 'Segoe UI', Arial, sans-serif;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 100vh;
    padding: 2rem;
  }}
  .header {{
    text-align: center;
    margin-bottom: 2rem;
  }}
  .header .practice-label {{
    display: inline-block;
    background: #f9a825;
    color: #3e2700;
    font-size: 1rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    padding: 0.3rem 1rem;
    border-radius: 4px;
    margin-bottom: 0.75rem;
  }}
  .header h1 {{
    font-size: 2.8rem;
    font-weight: 700;
    letter-spacing: 0.04em;
    text-transform: uppercase;
  }}
  .header p {{
    font-size: 1.3rem;
    opacity: 0.85;
    margin-top: 0.4rem;
  }}
  .content {{
    display: flex;
    gap: 4rem;
    align-items: flex-start;
    justify-content: center;
    flex-wrap: wrap;
    width: 100%;
  }}
  .candidates {{
    flex: 1;
    min-width: 320px;
    max-width: 520px;
  }}
  .candidates h2 {{
    font-size: 1.1rem;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    opacity: 0.7;
    margin-bottom: 0.75rem;
    color: #f9a825;
  }}
  .candidates ul {{
    list-style: none;
  }}
  .candidates li {{
    display: flex;
    align-items: center;
    gap: 1rem;
    font-size: 1.8rem;
    padding: 0.45rem 0;
    border-bottom: 1px solid rgba(249,168,37,0.2);
    font-style: italic;
    opacity: 0.9;
  }}
  .candidates li:last-child {{ border-bottom: none; }}
  .num {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 2.4rem;
    height: 2.4rem;
    border-radius: 50%;
    background: #f9a825;
    color: #3e2700;
    font-size: 1.1rem;
    font-weight: 700;
    font-style: normal;
    flex-shrink: 0;
  }}
  .qr-panel {{
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 1rem;
    min-width: 260px;
  }}
  .qr-panel h2 {{
    font-size: 1.1rem;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    opacity: 0.7;
    color: #f9a825;
  }}
  #qrcode canvas, #qrcode img {{ border-radius: 8px; }}
  .url-text {{
    font-size: 0.78rem;
    opacity: 0.55;
    word-break: break-all;
    max-width: 280px;
    text-align: center;
  }}
  .footer-note {{
    margin-top: 2.5rem;
    font-size: 1rem;
    opacity: 0.6;
    text-align: center;
    letter-spacing: 0.03em;
  }}
</style>
</head>
<body>
  <div class="header">
    <div class="practice-label">🗳️ Practice Session — Fictional Candidates Only</div>
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
  <p class="footer-note">
    Enter any number as your delegate number &nbsp;·&nbsp;
    Rank the candidates &nbsp;·&nbsp;
    No votes are recorded
  </p>

<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
<script>
  new QRCode(document.getElementById("qrcode"), {{
    text: "{url}",
    width: 240,
    height: 240,
    colorDark: "#000000",
    colorLight: "#ffffff",
    correctLevel: QRCode.CorrectLevel.M
  }});
</script>
</body>
</html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)


# ─────────────────────────────────────────────────────────────────────────
#  MAIN MENU
# ─────────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "═" * 50)
    print("  3rd Congressional District Convention")
    print("  Election Setup — Day-Of Operations Tool")
    print("  " + datetime.now().strftime("%A, %B %d, %Y  %I:%M %p"))
    print("═" * 50)

    if not QRCODE_AVAILABLE:
        print("\n  ⚠  NOTE: QR code PNG generation is unavailable.")
        print("     Install with:  pip install qrcode[pil]")
        print("     Projection displays will still include QR codes via JavaScript.\n")

    while True:
        print("\n╔══════════════════════════════════════╗")
        print("║           MAIN MENU                  ║")
        print("╚══════════════════════════════════════╝")
        print("  1. Manage nominees")
        print("  2. Manage delegate roster")
        print("  3. Record surrendered credentials")
        print("  4. Generate QR code & projection display")
        print("  5. Generate practice ballot QR code & display slide")
        print("  6. Exit")
        choice = input("\nChoice: ").strip()

        if choice == "1":
            manage_nominees()
        elif choice == "2":
            manage_delegates()
        elif choice == "3":
            manage_surrendered()
        elif choice == "4":
            generate_qr_and_display()
        elif choice == "5":
            generate_practice_qr()
        elif choice == "6":
            print("\n  Data saved. Goodbye.\n")
            sys.exit(0)
        else:
            print("  ⚠  Enter 1–6.")


if __name__ == "__main__":
    main()
