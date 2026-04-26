#!/usr/bin/env python3
"""
=============================================================================
3rd Congressional District Democratic Convention — May 2, 2026
ELECTION TABULATION TOOL  v2.0  —  Ranked-Choice IRV Tabulator
=============================================================================
Handles three elections:
  A. State Central Committee (SCC)   — Sections V.K / VI    (IRV, 4 seats each)
  B. DEI Committee Chair             — Sections V.K / IX    (IRV, 1 seat)
  C. State Convention Committee      — Sections V.K / XII   (Slate vote, 14 seats)

Ballot input:
  • Google Sheets CSV export   (primary — electronic ballots)
  • Paper ballot CSV           (optional — merged by delegate number)

Other inputs:
  • surrendered_delegates.json  (produced by district3_setup.py)

Outputs:
  • Round-by-round results on screen
  • Excel workbook  (Summary + one sheet per round per election)
  • Plain-text audit log  (for the official record)
  • Exceptions report  (Spoiled / Exhausted / Surrendered Credentials)

Requirements:  Python 3.8+,  openpyxl  (pip install openpyxl)
=============================================================================

SCC SEAT-RESTART RULE (confirmed by Rules Committee):
  After each SCC seat is won, the next seat's pool excludes BOTH the elected
  winner AND all candidates previously eliminated (below 15% threshold) in
  any prior seat's rounds.  Example:
    Round 1: A wins, C & D eliminated → Seat 2 pool: B, E, F, ...
    Round 2: B wins, E eliminated     → Seat 3 pool: F, ...
  Only candidates who were never elected and never eliminated continue forward.
=============================================================================
"""

import csv
import json
import os
import re
import sys
import datetime
from collections import defaultdict
from pathlib import Path

# ── Try to import openpyxl ──────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# =============================================================================
# CONSTANTS
# =============================================================================

VERSION           = "2.0.0"
CONVENTION_DATE   = "May 2, 2026"
DISTRICT          = "3rd Congressional District Democratic Convention"

ELIMINATION_THRESHOLD   = 0.15   # 15%  — candidates below this are eliminated
SCC_FIRST_BALLOT_MAX    = 0.50   # 50%  — max fraction of seats filled in Round 1
MAJORITY_THRESHOLD      = 0.50   # must receive MORE THAN 50% to be elected

# Column header patterns in the Google Sheets CSV export
# e.g. "Candidate A Rank", "Candidate B Rank"
RANK_HEADER_RE = re.compile(r'candidate\s+([A-Za-z])\s+rank', re.IGNORECASE)

# =============================================================================
# COLOR / DISPLAY HELPERS
# =============================================================================

class Color:
    RESET   = "\033[0m"
    BOLD    = "\033[1m"
    GREEN   = "\033[92m"
    YELLOW  = "\033[93m"
    RED     = "\033[91m"
    CYAN    = "\033[96m"
    MAGENTA = "\033[95m"
    WHITE   = "\033[97m"
    GREY    = "\033[90m"

def c(text, color):
    return f"{color}{text}{Color.RESET}"

def banner():
    w = 70
    print()
    print(c("=" * w, Color.CYAN))
    print(c(f"  {DISTRICT}".center(w), Color.BOLD + Color.WHITE))
    print(c(f"  Election Tabulation Tool  v{VERSION}".center(w), Color.CYAN))
    print(c(f"  {CONVENTION_DATE}".center(w), Color.GREY))
    print(c("=" * w, Color.CYAN))
    print()

def section_header(title):
    print()
    print(c("─" * 62, Color.CYAN))
    print(c(f"  {title}", Color.BOLD + Color.WHITE))
    print(c("─" * 62, Color.CYAN))
    log(f"\n{'─'*62}\n  {title}\n{'─'*62}", also_print=False)

def seat_ballot_header(seat_num, election_label):
    """Displayed once when a seat's tabulation begins."""
    msg = f"  SEAT {seat_num} BALLOT — {election_label}"
    print()
    print(c("╔" + "═" * 60 + "╗", Color.MAGENTA + Color.BOLD))
    print(c(f"║{msg:<60}║", Color.MAGENTA + Color.BOLD))
    print(c("╚" + "═" * 60 + "╝", Color.MAGENTA + Color.BOLD))
    log(f"\n{'═'*62}\n[SEAT {seat_num} BALLOT] {election_label}\n{'═'*62}", also_print=False)

def distribution_round_header(seat_round_num, seat_num, election_label):
    """Displayed for each redistribution round within a seat's ballot."""
    msg = f"  Seat {seat_num} Ballot — Distribution Round {seat_round_num}"
    print()
    print(c("┌" + "─" * 60 + "┐", Color.YELLOW))
    print(c(f"│{msg:<60}│", Color.YELLOW))
    print(c("└" + "─" * 60 + "┘", Color.YELLOW))
    log(f"\n[SEAT {seat_num} — Distribution Round {seat_round_num}] {election_label}", also_print=False)

def display_round_table(candidate_map, vote_counts, total_active, majority_needed,
                        elected=None, eliminated=None):
    """
    Print a formatted table of vote counts for one IRV round.
    candidate_map : {cand_num_str : name_str}
    vote_counts   : {cand_num_str : int}
    """
    elected    = elected    or []
    eliminated = eliminated or []

    threshold_votes = total_active * ELIMINATION_THRESHOLD

    sorted_cands = sorted(vote_counts.items(), key=lambda x: -x[1])

    col_w = [32, 8, 10, 16]
    header = (f"  {'Candidate':<{col_w[0]}} {'Votes':>{col_w[1]}} "
              f"{'  %':>{col_w[2]}} {'Status':<{col_w[3]}}")
    divider = "  " + "─" * (sum(col_w) + 3 * 2)

    print(c(header, Color.BOLD))
    print(c(divider, Color.GREY))
    log(header, also_print=False)
    log(divider, also_print=False)

    for num, votes in sorted_cands:
        name = candidate_map.get(num, f"Candidate {num}")
        pct  = (votes / total_active * 100) if total_active > 0 else 0

        if num in elected:
            status = "✓ ELECTED"
            color  = Color.GREEN
        elif num in eliminated:
            status = "✗ ELIMINATED"
            color  = Color.RED
        else:
            status = ""
            color  = Color.WHITE

        row = (f"  {name:<{col_w[0]}} {votes:>{col_w[1]}} "
               f"{pct:>{col_w[2]-1}.1f}%  {status:<{col_w[3]}}")
        print(c(row, color))
        log(row, also_print=False)

    summary = (f"\n  Active ballots   : {total_active}"
               f"\n  Majority needed  : {majority_needed}  (> 50% of active ballots)"
               f"\n  15% threshold    : {threshold_votes:.1f} votes")
    print(c(divider, Color.GREY))
    print(c(summary, Color.GREY))
    log(divider + summary, also_print=False)
    print()

def prompt(msg, default=None):
    suffix = f" [{default}]" if default is not None else ""
    return input(f"\n  {msg}{suffix}: ").strip() or (str(default) if default is not None else "")

def confirm(msg):
    ans = input(f"\n  {msg} (y/n): ").strip().lower()
    return ans in ("y", "yes")

def menu(title, options):
    print()
    print(c(f"  {title}", Color.BOLD))
    for i, opt in enumerate(options, 1):
        print(f"    {c(str(i), Color.CYAN)}.  {opt}")
    while True:
        try:
            choice = int(input("\n  Enter number: ").strip())
            if 1 <= choice <= len(options):
                return choice
        except ValueError:
            pass
        print(c("  Please enter a valid number.", Color.RED))

# =============================================================================
# AUDIT LOG
# =============================================================================

_audit_lines = []

def log(text, also_print=True):
    _audit_lines.append(text)
    if also_print:
        print(text)

def save_audit_log(filepath):
    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(_audit_lines))
    print(c(f"\n  Audit log saved → {filepath}", Color.GREEN))

# =============================================================================
# BALLOT DATA CLASS
# =============================================================================

class Ballot:
    """Represents one delegate's ranked-choice ballot."""
    __slots__ = ("delegate_number", "timestamp", "rankings",
                 "source", "is_spoiled", "spoil_reason")

    def __init__(self, delegate_number, timestamp, rankings, source="electronic"):
        self.delegate_number = str(delegate_number).strip()
        self.timestamp       = timestamp
        # rankings: {candidate_num_str: rank_int}
        # rank_int == 0 means "not ranked" (treat as unranked)
        self.rankings  = rankings
        self.source    = source        # "electronic" or "paper"
        self.is_spoiled  = False
        self.spoil_reason = ""

    def get_first_active_choice(self, active_keys):
        """
        Return the candidate letter of this ballot's highest-ranked (lowest rank int)
        candidate who is still active.  Returns None if ballot is exhausted.
        active_keys : iterable of candidate letter strings
        """
        best_num  = None
        best_rank = None
        for num in active_keys:
            rank = self.rankings.get(num)
            if rank and isinstance(rank, int) and rank > 0:
                if best_rank is None or rank < best_rank:
                    best_rank = rank
                    best_num  = num
        return best_num

# =============================================================================
# CSV LOADING — ELECTRONIC BALLOTS (Google Sheets export)
# =============================================================================
#
# Expected column headers (set by apps_script.js writeHeader):
#   Timestamp | Delegate Number | Election | Is Test | Is Auto-Submit |
#   Candidate A Rank | Candidate B Rank | Candidate C Rank | ...
#
# Each "Candidate A Rank" cell contains an integer (the rank the delegate
# assigned to that candidate) or is blank (candidate not ranked).

def load_electronic_ballots(filepath):
    """
    Load ranked-choice ballots from a Google Sheets CSV export.
    Returns (ballots: list[Ballot], candidate_keys: list[str], warnings: list[str])
    where candidate_keys are the letters found in the column headers.
    """
    ballots        = []
    warnings       = []
    candidate_keys = []

    try:
        with open(filepath, newline="", encoding="utf-8-sig") as f:
            reader     = csv.DictReader(f)
            fieldnames = [fn.strip() for fn in (reader.fieldnames or [])]

            # Discover candidate letter columns
            rank_cols = {}   # {candidate_letter_str: exact_header}
            for fn in fieldnames:
                m = RANK_HEADER_RE.match(fn.strip())
                if m:
                    rank_cols[m.group(1).upper()] = fn

            if not rank_cols:
                warnings.append("No 'Candidate A Rank' columns found in this CSV. "
                                 "Are you using the Google Sheets export?")
                return [], [], warnings

            candidate_keys = sorted(rank_cols.keys())

            for row_num, row in enumerate(reader, start=2):
                delegate_num = str(row.get("Delegate Number", row.get("delegate_number", ""))).strip()
                timestamp    = str(row.get("Timestamp", "")).strip()
                is_test      = str(row.get("Is Test", "NO")).strip().upper()

                if is_test == "YES":
                    continue   # skip test submissions

                # Parse rankings
                rankings = {}
                for cand_num, col in rank_cols.items():
                    raw = str(row.get(col, "")).strip()
                    if raw == "":
                        rankings[cand_num] = 0   # not ranked
                    else:
                        try:
                            rankings[cand_num] = int(raw)
                        except ValueError:
                            rankings[cand_num] = 0
                            warnings.append(f"Row {row_num}: non-integer rank '{raw}' "
                                            f"for Candidate {cand_num} (delegate {delegate_num}) — ignored")

                if not delegate_num:
                    warnings.append(f"Row {row_num}: missing delegate number — ballot skipped")
                    continue

                ballots.append(Ballot(delegate_num, timestamp, rankings, source="electronic"))

    except FileNotFoundError:
        warnings.append(f"File not found: {filepath}")

    return ballots, candidate_keys, warnings

# =============================================================================
# CSV LOADING — PAPER BALLOTS
# =============================================================================
#
# Paper ballot CSV uses the same format as the Google Sheets export:
#   Ballot_Number, Delegate_Number, Candidate_A_Rank, Candidate_B_Rank, ...
#
# "Candidate_A_Rank" columns contain the rank (1, 2, 3...) or blank.
# Column names are flexible — any column matching "Candidate A Rank" (case-
# insensitive) is treated as a rank column.

def load_paper_ballots(filepath):
    """
    Load paper ballots from CSV.
    Returns (ballots: list[Ballot], candidate_keys: list[str], warnings: list[str])
    """
    # Re-uses the same loader as electronic; paper format mirrors Google Sheets
    ballots, candidate_keys, warnings = load_electronic_ballots(filepath)
    for b in ballots:
        b.source = "paper"
    return ballots, candidate_keys, warnings

# =============================================================================
# BALLOT MERGE
# =============================================================================

def merge_ballots(electronic, paper):
    """
    Merge electronic and paper ballots.
    If a delegate appears in both, the electronic ballot wins
    (paper is kept only if delegate has no electronic ballot).
    Returns merged list[Ballot].
    """
    by_delegate = {}
    # Load electronic first (higher priority)
    for b in electronic:
        by_delegate[b.delegate_number] = b
    # Paper fills in anyone not found electronically
    paper_added = 0
    for b in paper:
        if b.delegate_number not in by_delegate:
            by_delegate[b.delegate_number] = b
            paper_added += 1
        else:
            log(f"  [MERGE] Delegate {b.delegate_number}: paper ballot superseded by electronic.", also_print=False)

    merged = list(by_delegate.values())
    log(f"\n[MERGE] Electronic: {len(electronic)}  Paper: {len(paper)}  "
        f"Paper used: {paper_added}  Total: {len(merged)}", also_print=False)
    return merged

# =============================================================================
# SURRENDERED CREDENTIALS CHECK
# =============================================================================

def check_surrendered(ballots, surrendered_path):
    """
    Load surrendered_delegates.json and flag any ballot submitted AFTER
    the delegate's surrender time as "Spoiled — Surrendered Credentials".

    surrendered_delegates.json format (created by district3_setup.py):
      [ { "delegate_number": "101", "name": "Jane Doe",
          "surrender_time": "2026-05-02 09:30:00" }, ... ]

    Returns (valid_ballots, spoiled_surrendered) lists.
    """
    surrendered_map = {}   # {delegate_number: surrender_datetime}

    if not os.path.isfile(surrendered_path):
        log("  [SURRENDERED] No surrendered_delegates.json found — skipping check.",
            also_print=False)
        return ballots, []

    try:
        with open(surrendered_path, encoding="utf-8") as f:
            records = json.load(f)
        for rec in records:
            dn = str(rec.get("delegate_number", "")).strip()
            ts = rec.get("surrender_time", "")
            if dn and ts:
                try:
                    # Accept both "YYYY-MM-DD HH:MM:SS" and ISO formats
                    dt = datetime.datetime.fromisoformat(ts)
                    surrendered_map[dn] = dt
                except ValueError:
                    pass
    except (json.JSONDecodeError, OSError) as e:
        log(f"  [WARNING] Could not read surrendered_delegates.json: {e}", also_print=False)
        return ballots, []

    valid   = []
    spoiled = []

    for b in ballots:
        if b.delegate_number in surrendered_map:
            surrender_dt = surrendered_map[b.delegate_number]
            # Try to parse ballot timestamp
            try:
                ballot_dt = datetime.datetime.fromisoformat(b.timestamp)
                if ballot_dt > surrender_dt:
                    b.is_spoiled   = True
                    b.spoil_reason = (f"Surrendered Credentials "
                                      f"(surrendered {surrender_dt.strftime('%H:%M')}, "
                                      f"voted {ballot_dt.strftime('%H:%M')})")
                    spoiled.append(b)
                    log(f"  [SPOILED] Delegate {b.delegate_number} — {b.spoil_reason}",
                        also_print=False)
                    continue
            except (ValueError, TypeError):
                # Can't parse timestamp — flag as possibly surrendered
                log(f"  [WARNING] Delegate {b.delegate_number} in surrendered list "
                    f"but ballot timestamp unreadable — included with caution.",
                    also_print=False)
        valid.append(b)

    return valid, spoiled

# =============================================================================
# CANDIDATE SETUP
# =============================================================================

def enter_candidates(election_name, found_keys=None):
    """
    Prompt operator to enter candidate names by letter.
    found_keys: list of candidate letter strings found in the CSV headers.
    Returns dict {letter_str: name_str}.
    """
    print()
    print(c(f"  Enter candidate names for: {election_name}", Color.BOLD))
    if found_keys:
        print(c(f"  (CSV has {len(found_keys)} candidate columns: "
                f"Candidates {', '.join(found_keys)})", Color.GREY))
        print(c("  Enter each candidate's name in the same lettered order.", Color.GREY))

    candidate_map = {}

    if found_keys:
        for key in found_keys:
            while True:
                name = input(f"    Name for Candidate {key}: ").strip()
                if name:
                    candidate_map[key] = name
                    break
                print(c("    Name cannot be blank.", Color.RED))
    else:
        # Manual entry (no CSV loaded yet)
        print(c("  Press Enter on a blank line when done.", Color.GREY))
        idx = 0
        while True:
            letter = chr(65 + idx)   # A, B, C, …
            name = input(f"    Candidate {letter} name (or Enter to finish): ").strip()
            if not name:
                if idx < 2:
                    print(c("  At least 2 candidates required.", Color.RED))
                    continue
                break
            candidate_map[letter] = name
            idx += 1

    log(f"\n[CANDIDATES — {election_name}]", also_print=False)
    for key, name in candidate_map.items():
        log(f"  {key}: {name}", also_print=False)

    print(c(f"\n  Confirmed {len(candidate_map)} candidates:", Color.CYAN))
    for key, name in candidate_map.items():
        print(f"    {key}. {name}")

    return candidate_map

def get_csv_filepath(label):
    """Prompt for a CSV path and validate it exists. Returns path or None."""
    print(c(f"\n  Enter path to CSV file for {label}:", Color.BOLD))
    print(c("  (You can drag the file into this window)", Color.GREY))
    while True:
        path = input("  Path: ").strip().strip('"').strip("'")
        if os.path.isfile(path):
            return path
        print(c(f"  File not found: {path}", Color.RED))
        if not confirm("  Try again?"):
            return None

def load_ballots_for_election(election_name, data_dir):
    """
    Interactive ballot loading:
    1. Ask for Google Sheets CSV (required unless paper-only).
    2. Optionally ask for paper ballot CSV.
    3. Merge; check surrendered credentials.
    Returns (ballots, candidate_map) or (None, None) on failure.
    """
    section_header(f"BALLOT LOADING — {election_name}")

    electronic_ballots = []
    paper_ballots      = []
    candidate_keys     = []
    warnings           = []

    # ── Electronic (Google Sheets) ───────────────────────────────────────────
    choice = menu("How will ballots be loaded?", [
        "Google Sheets CSV export  (recommended — electronic ballots)",
        "Paper ballot CSV only  (no electronic ballots)",
        "Both  (merge electronic + paper)",
    ])

    if choice in (1, 3):
        path = get_csv_filepath(f"{election_name} — Google Sheets export")
        if path:
            electronic_ballots, candidate_keys, warnings = load_electronic_ballots(path)
            for w in warnings:
                print(c(f"  ⚠  {w}", Color.YELLOW))
            print(c(f"\n  Loaded {len(electronic_ballots)} electronic ballot(s).", Color.GREEN))
            log(f"[LOAD] Electronic CSV: {path}  ({len(electronic_ballots)} ballots)", also_print=False)

    if choice in (2, 3):
        path = get_csv_filepath(f"{election_name} — paper ballots CSV")
        if path:
            paper_ballots, paper_keys, warnings = load_paper_ballots(path)
            for w in warnings:
                print(c(f"  ⚠  {w}", Color.YELLOW))
            print(c(f"\n  Loaded {len(paper_ballots)} paper ballot(s).", Color.GREEN))
            log(f"[LOAD] Paper CSV: {path}  ({len(paper_ballots)} ballots)", also_print=False)
            if not candidate_keys:
                candidate_keys = paper_keys

    if not electronic_ballots and not paper_ballots:
        print(c("\n  No ballots loaded — cannot tabulate.", Color.RED))
        return None, None

    # ── Merge ────────────────────────────────────────────────────────────────
    all_ballots = merge_ballots(electronic_ballots, paper_ballots)

    # ── Surrendered credentials ──────────────────────────────────────────────
    surrendered_path = os.path.join(data_dir, "surrendered_delegates.json")
    all_ballots, spoiled_surrendered = check_surrendered(all_ballots, surrendered_path)
    if spoiled_surrendered:
        print(c(f"\n  ⚠  {len(spoiled_surrendered)} ballot(s) flagged: Surrendered Credentials.", Color.YELLOW))
        for b in spoiled_surrendered:
            print(c(f"     Delegate {b.delegate_number} — {b.spoil_reason}", Color.YELLOW))

    # ── Candidate names ──────────────────────────────────────────────────────
    candidate_map = enter_candidates(election_name, candidate_keys)

    return all_ballots, candidate_map, spoiled_surrendered

# =============================================================================
# IRV CORE — VOTE COUNTING
# =============================================================================

def count_round_votes(ballots, active_nums):
    """
    Count first-active-choice votes for one IRV round.
    Returns:
      vote_counts   : {cand_num_str: int}   (only active candidates)
      total_active  : int  (ballots with an active choice — used for majority/threshold)
      exhausted     : int  (ballots with no active choice remaining)
    """
    vote_counts  = {num: 0 for num in active_nums}
    exhausted    = 0

    for b in ballots:
        if b.is_spoiled:
            continue
        choice = b.get_first_active_choice(active_nums)
        if choice is None:
            exhausted += 1
        else:
            vote_counts[choice] = vote_counts.get(choice, 0) + 1

    total_active = sum(vote_counts.values())
    return vote_counts, total_active, exhausted

def compute_elimination_cut(vote_counts, threshold=ELIMINATION_THRESHOLD):
    """
    Determine which candidates are eliminated based on the 15% rule,
    with the safeguard that no more than 49% of remaining candidates
    may be eliminated in a single round.

    Rules (Convention Rules Article V.K):
      1. Eliminate candidates with < 15% of total active votes.
      2. But if that would eliminate ≥ 50% of remaining candidates,
         lower the cut until fewer than 50% would be eliminated.
      3. Always eliminate at least the lowest vote-getter.

    Returns: (to_eliminate: list[str], actual_threshold_pct: float)
    """
    total = sum(vote_counts.values())
    if total == 0:
        return [], 0.0

    n = len(vote_counts)
    sorted_cands = sorted(vote_counts.items(), key=lambda x: x[1])

    # Step 1: Apply threshold
    cut_votes       = total * threshold
    below_threshold = [num for num, v in vote_counts.items() if v < cut_votes]
    actual_pct      = threshold * 100

    # Step 2: Safeguard — cannot eliminate ≥ 50% of candidates
    max_eliminate = max(1, (n - 1) // 2)

    if len(below_threshold) > max_eliminate:
        below_threshold = [num for num, v in sorted_cands[:max_eliminate]]
        if below_threshold:
            cut_votes  = vote_counts[below_threshold[-1]]
            actual_pct = cut_votes / total * 100

    # Step 3: Always eliminate at least the lowest
    lowest_num = sorted_cands[0][0]
    if lowest_num not in below_threshold:
        below_threshold.append(lowest_num)

    return below_threshold, actual_pct

# =============================================================================
# RUNOFF BALLOT GENERATION & TABULATION
# =============================================================================
#
# Triggered in Distribution Round 2+ when two or more candidates share the
# exact same lowest vote total and one must be dropped.
#
# Step 1 — Tiebreaker runoff ballot (the tied candidates only)
#           "Which candidate do you approve?  Select one."
#           Loser is eliminated; winner advances.
#
# Step 2 — Head-to-head ballot (if exactly 2 candidates remain after step 1)
#           Same question, now between the tiebreaker winner and the leader.
#           Winner of this ballot is elected to the seat.
#
# Both ballots use ballot.html with ?type=runoff in the URL.
# Results arrive as a separate Google Sheets CSV export (Runoff tab).

_ballot_base_url = ""   # set once per session when first runoff is needed

def _get_ballot_base_url():
    global _ballot_base_url
    if not _ballot_base_url:
        print(c("\n  A runoff ballot needs to be generated.", Color.CYAN))
        print(c("  Enter the GitHub Pages URL where ballot.html is hosted,", Color.GREY))
        print(c("  e.g.  https://yourname.github.io/ballot.html", Color.GREY))
        url = input("\n  Ballot URL: ").strip().rstrip("/")
        _ballot_base_url = url
    return _ballot_base_url

def _make_runoff_url(base_url, runoff_election_key, candidate_map_subset):
    """Build a ballot.html URL for a two-candidate runoff (type=runoff)."""
    cands_param = "-".join(
        f"{num}-{name.replace(' ', '%20')}"
        for num, name in candidate_map_subset.items()
    )
    return (f"{base_url}?election={runoff_election_key}"
            f"&candidates={cands_param}&seats=1&type=runoff")

def _show_runoff_qr(url, label):
    """Print the runoff URL and show a QR code if the qrcode library is available."""
    print(c(f"\n  ┌─ {label} ─────────────────────────────────────────────", Color.MAGENTA))
    print(c(f"  │  URL: {url}", Color.CYAN))
    print(c(f"  └────────────────────────────────────────────────────────", Color.MAGENTA))
    log(f"[RUNOFF URL — {label}]: {url}", also_print=False)
    try:
        import qrcode as _qr
        qr = _qr.QRCode(box_size=2, border=2)
        qr.add_data(url)
        qr.make(fit=True)
        qr.print_ascii(invert=True)
    except ImportError:
        print(c("  (Install qrcode[pil] to display QR code here, or use "
                "district3_setup.py to generate one.)", Color.GREY))

def _load_runoff_csv(filepath):
    """
    Load runoff ballot results from Google Sheets CSV export.
    Expected columns: Timestamp | Delegate Number | Election |
                      Is Test | Is Auto-Submit | Choice
    'Choice' contains the candidate LETTER the delegate selected.
    Returns ({num_str: count_int}, warnings_list).
    De-duplicates by delegate number (last submission wins).
    """
    by_delegate = {}
    warnings    = []
    try:
        with open(filepath, newline="", encoding="utf-8-sig") as f:
            reader     = csv.DictReader(f)
            fieldnames = [fn.strip() for fn in (reader.fieldnames or [])]
            choice_col = next(
                (fn for fn in fieldnames if fn.strip().lower() == "choice"), None)
            if not choice_col:
                warnings.append("No 'Choice' column found in the runoff CSV. "
                                 "Check that the Apps Script wrote the correct header.")
                return {}, warnings
            for row in reader:
                if str(row.get("Is Test", "NO")).strip().upper() == "YES":
                    continue
                delegate = str(row.get("Delegate Number", "")).strip()
                choice   = str(row.get(choice_col, "")).strip()
                if delegate and choice:
                    by_delegate[delegate] = choice
    except FileNotFoundError:
        warnings.append(f"File not found: {filepath}")
        return {}, warnings

    vote_counts = {}
    for choice in by_delegate.values():
        vote_counts[choice] = vote_counts.get(choice, 0) + 1
    return vote_counts, warnings

def run_runoff(runoff_type, candidate_map_subset, election_key,
               election_label, seat_num, seat_round_label, rounds_data_out):
    """
    Orchestrate one runoff ballot event.

    runoff_type         : "tiebreaker" | "head-to-head"
    candidate_map_subset: {num_str: name_str}  — exactly 2 entries
    seat_round_label    : string used in Excel tab, e.g. "Tiebreaker" or "H2H"
    rounds_data_out     : list to append the result record to

    Returns winner_num (str) or None on failure.
    """
    nums  = list(candidate_map_subset.keys())
    names = [candidate_map_subset[n] for n in nums]

    section_header(f"RUNOFF — {election_label} — Seat {seat_num} — "
                   f"{runoff_type.replace('-',' ').title()}")

    if runoff_type == "tiebreaker":
        print(c(f"\n  Tie for lowest between: {names[0]}  and  {names[1]}", Color.MAGENTA + Color.BOLD))
        print(c("  A tiebreaker runoff ballot will be issued. Delegates select ONE candidate.", Color.GREY))
    else:
        print(c(f"\n  Head-to-head ballot: {names[0]}  vs  {names[1]}", Color.MAGENTA + Color.BOLD))
        print(c("  Delegates select ONE candidate. The winner is elected to this seat.", Color.GREY))

    # Generate and display URL + QR code
    runoff_key = f"runoff-{election_key}-{''.join(nums)}"
    base_url   = _get_ballot_base_url()
    url        = _make_runoff_url(base_url, runoff_key, candidate_map_subset)
    _show_runoff_qr(url, f"Seat {seat_num} — {runoff_type.title()}")

    print(c("\n  ─── Chair instructions ───────────────────────────────────────", Color.GREY))
    print(c("  1. Project the QR code or read the URL aloud.", Color.GREY))
    print(c("  2. Announce to delegates:", Color.WHITE))
    print(c(f'     "Which candidate do you approve — {names[0]} or {names[1]}?'
            f' You may select only one."', Color.BOLD))
    print(c("  3. Allow time to vote, then close the ballot window.", Color.GREY))
    print(c("  4. Export the Runoff tab from Google Sheets as CSV.", Color.GREY))
    print(c("  ─────────────────────────────────────────────────────────────", Color.GREY))

    input(c("\n  Press Enter when ready to load the runoff CSV...", Color.YELLOW))
    path = get_csv_filepath(f"Runoff CSV — {runoff_type}")

    vote_counts = {}
    total       = 0
    winner_num  = None

    if path:
        vote_counts, warnings = _load_runoff_csv(path)
        for w in warnings:
            print(c(f"  ⚠  {w}", Color.YELLOW))

        if vote_counts:
            total = sum(vote_counts.values())
            print(c(f"\n  Results ({total} votes cast):", Color.BOLD))
            for num in sorted(nums, key=lambda n: -vote_counts.get(n, 0)):
                name = candidate_map_subset[num]
                cnt  = vote_counts.get(num, 0)
                pct  = cnt / total * 100 if total else 0
                bar  = "█" * int(pct / 5)
                print(c(f"    {name:<28} {cnt:>4}  ({pct:>5.1f}%)  {bar}", Color.WHITE))
                log(f"  [RUNOFF {runoff_type.upper()}] {name}: {cnt} ({pct:.1f}%)",
                    also_print=False)

            max_votes      = max(vote_counts.values())
            tied_in_runoff = [n for n, v in vote_counts.items() if v == max_votes]

            if len(tied_in_runoff) == 1:
                winner_num = tied_in_runoff[0]
            else:
                print(c("\n  ⚠  Runoff itself is tied. Chair must determine the winner.", Color.RED))
                log("[RUNOFF TIE] Runoff result tied — Chair determination required.", also_print=False)
                print(c("  Enter the winner:", Color.YELLOW))
                for i, n in enumerate(tied_in_runoff, 1):
                    print(f"    {i}. {candidate_map_subset[n]}")
                while True:
                    try:
                        ch = int(input("  Enter number: ").strip())
                        if 1 <= ch <= len(tied_in_runoff):
                            winner_num = tied_in_runoff[ch - 1]
                            break
                    except ValueError:
                        pass
                    print(c("  Please enter a valid number.", Color.RED))
    else:
        # No CSV — manual entry fallback
        print(c("\n  No CSV loaded. Enter the winner manually:", Color.YELLOW))
        for i, n in enumerate(nums, 1):
            print(f"    {i}. {candidate_map_subset[n]}")
        while True:
            try:
                ch = int(input("  Enter number: ").strip())
                if 1 <= ch <= len(nums):
                    winner_num = nums[ch - 1]
                    break
            except ValueError:
                pass
            print(c("  Please enter a valid number.", Color.RED))
        vote_counts = {n: 0 for n in nums}

    if winner_num:
        loser_nums  = [n for n in nums if n != winner_num]
        winner_name = candidate_map_subset[winner_num]
        loser_names = [candidate_map_subset[n] for n in loser_nums]

        if runoff_type == "head-to-head":
            print(c(f"\n  ✓ {winner_name} ELECTED (Seat {seat_num} — head-to-head runoff)",
                    Color.GREEN + Color.BOLD))
        else:
            print(c(f"\n  {winner_name} advances.  "
                    f"{', '.join(loser_names)} eliminated.", Color.CYAN))

        log(f"[RUNOFF {runoff_type.upper()} — Seat {seat_num}] "
            f"Winner: {winner_name}  Eliminated: {loser_names}", also_print=False)

        rounds_data_out.append({
            "seat_num":        seat_num,
            "seat_round_num":  seat_round_label,
            "vote_counts":     {n: vote_counts.get(n, 0) for n in nums},
            "total_active":    total,
            "exhausted":       0,
            "majority_needed": total // 2 + 1 if total else 1,
            "elected":         [winner_num] if runoff_type == "head-to-head" else [],
            "eliminated":      loser_nums,
            "elim_reason":     f"{runoff_type} runoff ballot",
        })

    return winner_num

# =============================================================================
# ELECTION A — STATE CENTRAL COMMITTEE (SCC)   Article VI / V.K
# =============================================================================

def run_irv_one_seat(ballots, candidate_map, active_nums, election_label,
                     seat_num, max_wins_this_seat, rounds_data_out,
                     election_key=""):
    """
    Run IRV for a single seat among the given active_nums.
    Appends round data dicts to rounds_data_out (one per redistribution round).
    Each dict includes 'seat_round_num' (1, 2, 3… within this seat).
    Returns winner_num (str) or None.
    """
    active       = list(active_nums)
    seat_round   = 1   # redistribution round counter, resets to 1 for each seat

    # Show the seat ballot header once before redistribution begins
    seat_ballot_header(seat_num, election_label)

    while active:
        distribution_round_header(seat_round, seat_num, election_label)

        vote_counts, total_active, exhausted = count_round_votes(ballots, active)
        majority_needed = int(total_active * MAJORITY_THRESHOLD) + 1

        # Potential winners (> 50%)
        potential_winners = [num for num, v in vote_counts.items()
                             if v >= majority_needed]

        # Cap winners for the first seat's first redistribution round
        if seat_round == 1 and seat_num == 1:
            potential_winners = potential_winners[:max_wins_this_seat]

        winner_num = potential_winners[0] if potential_winners else None

        if winner_num:
            display_round_table(candidate_map, vote_counts, total_active, majority_needed,
                                elected=[winner_num])
            rounds_data_out.append({
                "seat_num":      seat_num,
                "seat_round_num": seat_round,
                "vote_counts":   dict(vote_counts),
                "total_active":  total_active,
                "exhausted":     exhausted,
                "majority_needed": majority_needed,
                "elected":       [winner_num],
                "eliminated":    [],
            })
            name = candidate_map.get(winner_num, winner_num)
            print(c(f"\n  ✓ {name} ELECTED (Seat {seat_num})", Color.GREEN + Color.BOLD))
            log(f"[SEAT {seat_num} — Round {seat_round}] WINNER: {name}", also_print=False)
            return winner_num

        # No winner — determine eliminations
        # Round 1: apply 15% threshold (may eliminate multiple candidates)
        # Round 2+: only the single lowest vote-getter is dropped
        sorted_by_votes = sorted(vote_counts.items(), key=lambda x: x[1])

        if seat_round == 1:
            to_eliminate, eff_pct = compute_elimination_cut(vote_counts)
            elim_reason = f"below {eff_pct:.1f}% elimination threshold"
        else:
            # Round 2+: check for a tie before eliminating
            lowest_votes    = sorted_by_votes[0][1]
            tied_for_lowest = [num for num, v in vote_counts.items() if v == lowest_votes]

            if len(tied_for_lowest) == 1:
                # Clear lowest — normal single-candidate elimination
                to_eliminate = tied_for_lowest
                lowest_pct   = lowest_votes / total_active * 100 if total_active > 0 else 0
                elim_reason  = f"lowest vote-getter ({lowest_pct:.1f}%)"

            else:
                # ── Tie for lowest — runoff ballot required ─────────────────
                tied_names = [candidate_map.get(n, n) for n in tied_for_lowest]
                print(c(f"\n  Tie for lowest ({', '.join(tied_names)}) "
                        f"— runoff ballot required.", Color.YELLOW + Color.BOLD))
                log(f"[SEAT {seat_num} — Round {seat_round}] "
                    f"TIE for lowest: {tied_names}", also_print=False)

                # Show this round's vote table with all tied-lowest highlighted
                display_round_table(candidate_map, vote_counts, total_active, majority_needed,
                                    eliminated=tied_for_lowest)
                rounds_data_out.append({
                    "seat_num":        seat_num,
                    "seat_round_num":  seat_round,
                    "vote_counts":     dict(vote_counts),
                    "total_active":    total_active,
                    "exhausted":       exhausted,
                    "majority_needed": majority_needed,
                    "elected":         [],
                    "eliminated":      list(tied_for_lowest),
                    "elim_reason":     "tie for lowest — tiebreaker runoff ballot issued",
                })

                # Step 1: Tiebreaker runoff among the tied candidates
                tied_map  = {n: candidate_map[n] for n in tied_for_lowest}
                tb_winner = run_runoff(
                    "tiebreaker", tied_map, election_key, election_label,
                    seat_num, "Tiebreaker", rounds_data_out,
                )
                if tb_winner is None:
                    return None   # operator cancelled — abort this seat

                # Remove tiebreaker loser(s) from the active pool
                to_elim_tb = [n for n in tied_for_lowest if n != tb_winner]
                for num in to_elim_tb:
                    active.remove(num)
                    print(c(f"  ✗ {candidate_map.get(num, num)} eliminated  "
                            f"(tiebreaker runoff)", Color.RED))
                    log(f"[SEAT {seat_num}] Tiebreaker loser: "
                        f"{candidate_map.get(num, num)}", also_print=False)

                # Step 2: If exactly 2 candidates now remain, run head-to-head ballot
                if len(active) == 2:
                    h2h_map    = {n: candidate_map[n] for n in active}
                    h2h_winner = run_runoff(
                        "head-to-head", h2h_map, election_key, election_label,
                        seat_num, "H2H", rounds_data_out,
                    )
                    if h2h_winner:
                        return h2h_winner

                # More than 2 remain (or head-to-head failed) — continue IRV
                seat_round += 1
                continue   # skip the normal display/append/eliminate block below

        display_round_table(candidate_map, vote_counts, total_active, majority_needed,
                            eliminated=to_eliminate)

        rounds_data_out.append({
            "seat_num":        seat_num,
            "seat_round_num":  seat_round,
            "vote_counts":     dict(vote_counts),
            "total_active":    total_active,
            "exhausted":       exhausted,
            "majority_needed": majority_needed,
            "elected":         [],
            "eliminated":      list(to_eliminate),
            "elim_reason":     elim_reason,
        })
        elim_names = [candidate_map.get(n, n) for n in to_eliminate]
        log(f"[SEAT {seat_num} — Round {seat_round}] No winner.  "
            f"Eliminated: {elim_names} ({elim_reason})", also_print=False)

        for num in to_eliminate:
            name = candidate_map.get(num, num)
            active.remove(num)
            print(c(f"  ✗ {name} eliminated  ({elim_reason})", Color.RED))

        if len(active) == 1:
            # Last candidate standing — elect them
            winner_num = active[0]
            name = candidate_map.get(winner_num, winner_num)
            print(c(f"\n  Only one candidate remaining — {name} is elected.", Color.YELLOW))
            print(c(f"  ✓ {name} ELECTED (Seat {seat_num})", Color.GREEN + Color.BOLD))
            log(f"[SEAT {seat_num}] Last candidate standing: {name}", also_print=False)
            rounds_data_out.append({
                "seat_num":       seat_num,
                "seat_round_num": seat_round + 1,
                "vote_counts":    {winner_num: total_active},
                "total_active":   total_active,
                "exhausted":      0,
                "majority_needed": majority_needed,
                "elected":        [winner_num],
                "eliminated":     [],
            })
            return winner_num

        seat_round += 1

    return None

def run_scc_election(gender_label, seats, ballots, candidate_map,
                     spoiled_surrendered, output_dir):
    """
    Run one gender category of the SCC election (sequential per-seat IRV).

    SEAT RESTART LOGIC (see note at top of file):
      After each winner, ALL non-elected candidates return to the pool.
      Only elected candidates are removed.
    """
    section_header(f"SCC — {gender_label.upper()} ({seats} seats)")
    log(f"\n[ELECTION] SCC — {gender_label} — {seats} seats", also_print=False)
    log(f"[DATE] {datetime.datetime.now().isoformat()}", also_print=False)

    all_nums       = list(candidate_map.keys())
    elected_all    = []       # candidate nums elected so far (in order)
    eliminated_all = set()    # candidate nums eliminated in ANY prior seat's rounds
    rounds_data    = []       # all round records across all seats

    # election_key is used when generating runoff ballot URLs
    election_key = "scc-w" if "women" in gender_label.lower() else "scc-mnb"

    # Seat 1 ballot cap: no more than floor(seats * 0.5) can be elected
    # in the very first redistribution round of Seat 1.
    max_seat1_round1_wins = max(1, int(seats * SCC_FIRST_BALLOT_MAX))
    print(c(f"\n  Seat 1 Ballot rule: no more than {max_seat1_round1_wins} of "
            f"{seats} seats may be filled in the first distribution round.", Color.CYAN))
    log(f"[RULE] Seat 1 first-distribution-round max wins: {max_seat1_round1_wins}",
        also_print=False)

    # Ratification path: nominees == seats
    if len(all_nums) == seats:
        print(c(f"\n  {len(all_nums)} nominees for {seats} seats — "
                f"proceeding to ratification ballot.", Color.YELLOW))
        vote_counts, total_active, exhausted = count_round_votes(ballots, all_nums)
        majority_needed = int(total_active * MAJORITY_THRESHOLD) + 1
        elected = [num for num, v in vote_counts.items() if v >= majority_needed]
        display_round_table(candidate_map, vote_counts, total_active, majority_needed,
                            elected=elected)
        rounds_data.append({
            "round_num": 1, "seat_num": 0,
            "vote_counts": vote_counts, "total_active": total_active,
            "exhausted": exhausted, "majority_needed": majority_needed,
            "elected": elected, "eliminated": [],
        })
        log(f"[RESULT] Ratification — Elected: {[candidate_map.get(n,n) for n in elected]}", also_print=False)
        if rounds_data:
            _save_scc_excel(gender_label, candidate_map, rounds_data, output_dir)
        return elected, rounds_data, spoiled_surrendered

    # Sequential per-seat IRV
    for seat in range(1, seats + 1):
        if not candidate_map:
            print(c("\n  No candidates remaining.", Color.YELLOW))
            break

        # SEAT RESTART: pool = all original candidates minus elected AND minus
        # any previously eliminated in prior seat rounds (eliminations carry forward).
        active_pool = [num for num in all_nums
                       if num not in elected_all and num not in eliminated_all]

        if len(active_pool) == 0:
            print(c(f"\n  No remaining candidates for Seat {seat}.", Color.YELLOW))
            break
        if len(active_pool) == 1:
            # Only one left — elect them
            winner_num = active_pool[0]
            name = candidate_map.get(winner_num, winner_num)
            elected_all.append(winner_num)
            print(c(f"\n  Only one candidate remaining — {name} ELECTED (Seat {seat}).", Color.GREEN + Color.BOLD))
            log(f"[SEAT {seat}] Uncontested: {name}", also_print=False)
            continue

        prev_rounds_count = len(rounds_data)

        winner_num = run_irv_one_seat(
            ballots, candidate_map, active_pool,
            f"SCC {gender_label}", seat,
            max_seat1_round1_wins if seat == 1 else 1,
            rounds_data,
            election_key=election_key,
        )

        if winner_num:
            elected_all.append(winner_num)
            # Collect eliminations from this seat's rounds and carry them forward
            this_seat_rounds = rounds_data[prev_rounds_count:]
            for rd in this_seat_rounds:
                for elim_num in rd["eliminated"]:
                    eliminated_all.add(elim_num)

            if eliminated_all:
                elim_names = [candidate_map.get(n, n) for n in eliminated_all
                              if n not in elected_all]
                print(c(f"  (Permanently eliminated — will not appear on next seat ballot: "
                        f"{', '.join(elim_names)})", Color.GREY))
        else:
            print(c(f"\n  WARNING: No winner found for Seat {seat}.", Color.RED))

    # ── Final results ────────────────────────────────────────────────────────
    print()
    print(c(f"  FINAL RESULT — SCC {gender_label} ({seats} seats)", Color.BOLD + Color.WHITE))
    for i, num in enumerate(elected_all, 1):
        print(c(f"    {i}. {candidate_map.get(num, num)}", Color.GREEN + Color.BOLD))
    log(f"\n[FINAL] SCC {gender_label} — Elected: "
        f"{[candidate_map.get(n,n) for n in elected_all]}", also_print=False)

    # Print exceptions
    _print_exceptions(ballots, spoiled_surrendered, election_name=f"SCC {gender_label}")

    # Save Excel
    if rounds_data:
        _save_scc_excel(gender_label, candidate_map, rounds_data, output_dir)

    return elected_all, rounds_data, spoiled_surrendered

def _save_scc_excel(gender_label, candidate_map, rounds_data, output_dir):
    safe = gender_label.replace("/", "_").replace(" ", "_")
    path = os.path.join(output_dir, f"SCC_{safe}_Results.xlsx")
    create_excel_report(f"SCC — {gender_label}", candidate_map, rounds_data, path)

def run_scc(ballots_women, candidate_map_women, spoiled_women,
            ballots_mnb,   candidate_map_mnb,   spoiled_mnb,
            output_dir):
    """Run the full SCC election (Women + Men/Non-Binary)."""
    section_header("STATE CENTRAL COMMITTEE (SCC) ELECTION — Article VI")
    print(c("\n  Two separate ballots: Women (4 seats) and Men/Non-Binary (4 seats).\n", Color.CYAN))

    results = {}

    elected_w, _, _ = run_scc_election(
        "Women", 4, ballots_women, candidate_map_women, spoiled_women, output_dir)
    results["Women"] = elected_w

    if confirm("\n  Proceed to Men/Non-Binary ballot?"):
        elected_mnb, _, _ = run_scc_election(
            "Men_NonBinary", 4, ballots_mnb, candidate_map_mnb, spoiled_mnb, output_dir)
        results["Men_NonBinary"] = elected_mnb

    return results

# =============================================================================
# ELECTION B — DEI COMMITTEE CHAIR   Article IX / V.K
# =============================================================================

def run_dei_election(ballots, candidate_map, spoiled_surrendered, output_dir):
    """
    Run the DEI Committee Chair election (1 seat, IRV with special tie rule).

    Tie rule (Convention Rules V.K):
      If two or more candidates are tied for the lowest total, eliminate
      all tied candidates UNLESS doing so would leave fewer than 2 —
      in that case, hold a special tiebreak round among only the tied candidates.
    """
    section_header("DEI COMMITTEE CHAIR ELECTION — Article IX")
    log("\n[ELECTION] DEI Committee Chair — 1 seat", also_print=False)
    log(f"[DATE] {datetime.datetime.now().isoformat()}", also_print=False)

    all_nums    = list(candidate_map.keys())
    active      = list(all_nums)
    rounds_data = []
    seat_round  = 1    # redistribution round within this single-seat election
    winner_num  = None

    # DEI has one seat — show the seat ballot header once
    seat_ballot_header(1, "DEI Committee Chair")

    while not winner_num and active:
        distribution_round_header(seat_round, 1, "DEI Committee Chair")

        vote_counts, total_active, exhausted = count_round_votes(ballots, active)
        majority_needed = int(total_active * MAJORITY_THRESHOLD) + 1

        # Check for winner
        best_num = max(vote_counts, key=vote_counts.get)
        if vote_counts[best_num] >= majority_needed:
            winner_num = best_num
            display_round_table(candidate_map, vote_counts, total_active, majority_needed,
                                elected=[winner_num])
            rounds_data.append({
                "seat_num":       1,
                "seat_round_num": seat_round,
                "vote_counts":    dict(vote_counts), "total_active": total_active,
                "exhausted":      exhausted, "majority_needed": majority_needed,
                "elected":        [winner_num], "eliminated": [],
            })
            name = candidate_map.get(winner_num, winner_num)
            print(c(f"\n  ✓ {name} ELECTED as DEI Committee Chair", Color.GREEN + Color.BOLD))
            log(f"[SEAT 1 — Round {seat_round}] WINNER: {name}", also_print=False)
            break

        # No winner — apply elimination rule:
        #   Distribution Round 1 : 15% threshold (may drop multiple), leave ≥ 2
        #   Distribution Round 2+: only the single lowest vote-getter drops
        # Tie rule (both rounds): if multiple candidates share the lowest total and
        # eliminating all would leave < 2, hold a special tiebreak distribution.
        sorted_cands = sorted(vote_counts.items(), key=lambda x: x[1])
        lowest_votes = sorted_cands[0][1]
        tied_lowest  = [num for num, v in vote_counts.items() if v == lowest_votes]
        to_eliminate = []
        elim_reason  = ""

        if seat_round == 1:
            # ── Round 1: 15% threshold ────────────────────────────────────────
            to_eliminate, eff_pct = compute_elimination_cut(vote_counts)
            # Guarantee ≥ 2 survivors
            survivors = [n for n in active if n not in to_eliminate]
            while len(survivors) < 2 and len(to_eliminate) > 1:
                by_votes     = sorted([(n, vote_counts[n]) for n in to_eliminate], key=lambda x: x[1])
                to_eliminate = [n for n, _ in by_votes[:-1]]
                survivors    = [n for n in active if n not in to_eliminate]
            if not to_eliminate:
                to_eliminate = [sorted_cands[0][0]]
            elim_reason = f"below {eff_pct:.1f}% elimination threshold (Round 1 rule)"

        else:
            # ── Round 2+: lowest drops ────────────────────────────────────────
            if len(tied_lowest) > 1:
                # Multiple candidates share the lowest total
                after_elim = len(active) - len(tied_lowest)
                if after_elim >= 2:
                    # Safe to eliminate all tied-lowest
                    to_eliminate = tied_lowest
                    names = [candidate_map.get(n, n) for n in tied_lowest]
                    print(c(f"\n  Tie for lowest ({', '.join(names)}) — all eliminated.", Color.YELLOW))
                    elim_reason = "tied for lowest vote-getter"
                else:
                    # Eliminating all would leave < 2 — special tiebreak distribution
                    names = [candidate_map.get(n, n) for n in tied_lowest]
                    print(c(f"\n  Tie for lowest: eliminating all would leave < 2 candidates.", Color.YELLOW))
                    print(c(f"  Special tiebreak distribution among: {', '.join(names)}", Color.YELLOW))
                    log(f"[SEAT 1 — Round {seat_round}] TIEBREAK among: {names}", also_print=False)

                    tie_votes, tie_total, _ = count_round_votes(ballots, tied_lowest)
                    tie_sorted    = sorted(tie_votes.items(), key=lambda x: x[1])
                    lowest_in_tie = tie_sorted[0][0]
                    to_eliminate  = [lowest_in_tie]
                    name          = candidate_map.get(lowest_in_tie, lowest_in_tie)
                    print(c(f"  Tiebreak result: {name} eliminated.", Color.RED))
                    log(f"[TIEBREAK] Eliminated: {name}", also_print=False)
                    elim_reason = "lowest in tiebreak distribution"
            else:
                # Clear lowest — eliminate them
                to_eliminate = [sorted_cands[0][0]]
                lowest_pct   = lowest_votes / total_active * 100 if total_active > 0 else 0
                elim_reason  = f"lowest vote-getter ({lowest_pct:.1f}%)"

        display_round_table(candidate_map, vote_counts, total_active, majority_needed,
                            eliminated=to_eliminate)

        rounds_data.append({
            "seat_num":        1,
            "seat_round_num":  seat_round,
            "vote_counts":     dict(vote_counts), "total_active": total_active,
            "exhausted":       exhausted, "majority_needed": majority_needed,
            "elected":         [], "eliminated": list(to_eliminate),
            "elim_reason":     elim_reason,
        })
        log(f"[SEAT 1 — Round {seat_round}] Eliminated: "
            f"{[candidate_map.get(n,n) for n in to_eliminate]} — {elim_reason}", also_print=False)

        for num in to_eliminate:
            name = candidate_map.get(num, num)
            active.remove(num)
            print(c(f"  ✗ {name} eliminated  ({elim_reason})", Color.RED))

        if len(active) == 1:
            winner_num = active[0]
            name = candidate_map.get(winner_num, winner_num)
            print(c(f"\n  Only one candidate remaining — {name} is elected.", Color.YELLOW))
            print(c(f"  ✓ {name} ELECTED as DEI Committee Chair", Color.GREEN + Color.BOLD))
            log(f"[SEAT 1 — Round {seat_round}] Last remaining: {name}", also_print=False)

        seat_round += 1

    # Exceptions + Excel
    _print_exceptions(ballots, spoiled_surrendered, election_name="DEI Committee Chair")
    if rounds_data:
        path = os.path.join(output_dir, "DEI_Chair_Results.xlsx")
        create_excel_report("DEI Committee Chair", candidate_map, rounds_data, path)

    if winner_num:
        log(f"\n[FINAL] DEI Chair — Elected: {candidate_map.get(winner_num, winner_num)}",
            also_print=False)
    return winner_num

# =============================================================================
# ELECTION C — STATE CONVENTION COMMITTEE (14 seats)   Article XII
# =============================================================================
#
# Slate vote: delegates vote YES or NO on the full slate.
# If slate approved → all nominees elected.
# If slate fails → individual IRV: 14 highest vote-getters elected.
# CSV column: "Slate Vote" containing YES / NO

def run_committee_election(output_dir):
    """Run the State Convention Committee election (slate vote, 14 seats)."""
    section_header("STATE CONVENTION COMMITTEE ELECTION — Article XII")
    log("\n[ELECTION] State Convention Committee — 14 seats (slate vote)", also_print=False)
    log(f"[DATE] {datetime.datetime.now().isoformat()}", also_print=False)

    seats = 14
    print(c("\n  This is a slate vote (YES / NO on the full slate).", Color.CYAN))
    print(c("  Nominees should already be recorded by the Committee.", Color.GREY))

    # Load CSV
    choice = menu("How will slate votes be loaded?", [
        "Google Sheets CSV export",
        "Paper ballot CSV  (Slate Vote column: YES or NO)",
        "Enter totals manually",
    ])

    yes_votes = 0
    no_votes  = 0
    nominee_names = []

    if choice in (1, 2):
        label = "Convention Committee — Slate Vote"
        path  = get_csv_filepath(label)
        if path:
            try:
                with open(path, newline="", encoding="utf-8-sig") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        vote = str(row.get("Slate Vote", row.get("slate_vote", ""))).strip().upper()
                        if str(row.get("Is Test","NO")).strip().upper() == "YES":
                            continue
                        if vote in ("YES", "Y", "1"):
                            yes_votes += 1
                        elif vote in ("NO", "N", "0"):
                            no_votes  += 1
                log(f"[LOAD] Slate CSV: {path}", also_print=False)
            except FileNotFoundError:
                print(c(f"  File not found: {path}", Color.RED))
    else:
        while True:
            try:
                yes_votes = int(input("\n  Total YES votes: ").strip())
                no_votes  = int(input("  Total NO votes:  ").strip())
                break
            except ValueError:
                print(c("  Please enter whole numbers.", Color.RED))

    total = yes_votes + no_votes
    print(c(f"\n  YES: {yes_votes}   NO: {no_votes}   Total: {total}", Color.CYAN))
    log(f"\n[SLATE VOTE] YES: {yes_votes}  NO: {no_votes}  Total: {total}", also_print=False)

    majority_needed = int(total * MAJORITY_THRESHOLD) + 1
    slate_approved  = yes_votes >= majority_needed

    if slate_approved:
        print(c(f"\n  Slate APPROVED ({yes_votes}/{total} = {yes_votes/total*100:.1f}%)", Color.GREEN + Color.BOLD))
        print(c("  All nominated slate members are elected.", Color.GREEN))
        log("[RESULT] Slate approved — all nominees elected.", also_print=False)
    else:
        print(c(f"\n  Slate FAILED ({yes_votes}/{total} = {yes_votes/total*100:.1f}%)", Color.RED + Color.BOLD))
        print(c("  Per Article XII.F.2: individual vote — 14 highest vote-getters elected.", Color.YELLOW))
        log("[RESULT] Slate failed — individual vote required.", also_print=False)
        print(c("\n  Individual vote not yet supported in this version.\n"
                "  Please run the election using the aggregate method and record results.", Color.YELLOW))

    # Save simple Excel summary
    if EXCEL_AVAILABLE:
        path = os.path.join(output_dir, "Convention_Committee_Slate_Vote.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Slate Vote"
        hdr_fill = PatternFill("solid", fgColor="1F3864")
        hdr_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        headers  = ["Vote Type", "Count", "% of Total", "Result"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = hdr_fill;  cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center")
        for row_i, (label, count) in enumerate([("YES", yes_votes), ("NO", no_votes)], 2):
            pct  = count / total * 100 if total else 0
            vals = [label, count, f"{pct:.1f}%",
                    ("APPROVED" if slate_approved else "FAILED") if label == "YES" else ""]
            for col, val in enumerate(vals, 1):
                ws.cell(row=row_i, column=col, value=val)
        ws.column_dimensions["A"].width = 14
        for ch in ["B","C","D"]:
            ws.column_dimensions[ch].width = 18
        wb.save(path)
        print(c(f"\n  Excel saved → {path}", Color.GREEN))

    return slate_approved

# =============================================================================
# EXCEPTIONS REPORT
# =============================================================================

def _print_exceptions(ballots, spoiled_surrendered, election_name=""):
    """Print a summary of spoiled, exhausted, and surrendered credential ballots."""
    exhausted = sum(
        1 for b in ballots
        if not b.is_spoiled and b.get_first_active_choice(list(b.rankings.keys())) is None
    )
    n_spoiled     = sum(1 for b in ballots if b.is_spoiled)
    n_surrendered = len(spoiled_surrendered)

    print()
    print(c("  ─── EXCEPTIONS REPORT ───────────────────────────────────", Color.GREY))
    print(c(f"  Election : {election_name}", Color.GREY))
    print(c(f"  Spoiled / Invalid ballots   : {n_spoiled}", Color.YELLOW if n_spoiled else Color.GREY))
    print(c(f"  Surrendered Credentials     : {n_surrendered}", Color.YELLOW if n_surrendered else Color.GREY))
    print(c(f"  Exhausted (blank/no choice) : {exhausted}", Color.GREY))
    print(c("  ─────────────────────────────────────────────────────────", Color.GREY))
    print()

    log(f"\n[EXCEPTIONS — {election_name}]", also_print=False)
    log(f"  Spoiled/Invalid   : {n_spoiled}", also_print=False)
    log(f"  Surrendered Creds : {n_surrendered}", also_print=False)
    log(f"  Exhausted         : {exhausted}", also_print=False)
    for b in spoiled_surrendered:
        log(f"  SURRENDERED: Delegate {b.delegate_number} — {b.spoil_reason}", also_print=False)

# =============================================================================
# EXCEL EXPORT
# =============================================================================

def create_excel_report(election_name, candidate_map, rounds_data, output_path):
    """
    Write an Excel workbook: Summary sheet + one sheet per redistribution round.
    Each redistribution round gets its own sheet so the record shows exactly
    why candidates were dropped between rounds.

    rounds_data: list of dicts with keys:
      seat_num, seat_round_num, vote_counts, total_active, exhausted,
      majority_needed, elected, eliminated
    """
    if not EXCEL_AVAILABLE:
        print(c("\n  openpyxl not installed — skipping Excel export.", Color.YELLOW))
        print(c("  Run: pip install openpyxl", Color.YELLOW))
        return

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    hdr_fill   = PatternFill("solid", fgColor="1F3864")
    elect_fill = PatternFill("solid", fgColor="C6EFCE")
    elim_fill  = PatternFill("solid", fgColor="FFCCCC")
    exh_fill   = PatternFill("solid", fgColor="F2F2F2")
    hdr_font   = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    bold_font  = Font(name="Calibri", bold=True, size=11)
    norm_font  = Font(name="Calibri", size=11)
    thin       = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"),  bottom=Side(style="thin"))

    def sc(cell, fill=None, font=None, align="left"):
        if fill:  cell.fill = fill
        cell.font      = font or norm_font
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border    = thin

    # ── One sheet per redistribution round ──────────────────────────────────
    for rd in rounds_data:
        seat_n   = rd.get("seat_num", "")
        seat_rn  = rd.get("seat_round_num", 1)

        # Sheet tab: "Seat 1 - Round 1", "Seat 1 - Tiebreaker", "Seat 1 - H2H" …
        # seat_rn may be an int (normal round) or a string (runoff label)
        seat_rn_str = str(seat_rn)
        if seat_n:
            if isinstance(seat_rn, int):
                tab_title  = f"Seat {seat_n} - Round {seat_rn}"
                full_title = f"Seat {seat_n} Ballot — Distribution Round {seat_rn}"
            else:
                tab_title  = f"Seat {seat_n} - {seat_rn_str}"
                full_title = f"Seat {seat_n} Ballot — {seat_rn_str} Runoff"
        else:
            tab_title  = f"Round {seat_rn_str}"
            full_title = f"Distribution Round {seat_rn_str}"

        ws = wb.create_sheet(title=tab_title[:31])   # Excel 31-char limit

        ws.merge_cells("A1:F1")
        ws["A1"]           = f"{election_name} — {full_title}"
        ws["A1"].font      = Font(name="Calibri", bold=True, size=13, color="1F3864")
        ws["A1"].alignment = Alignment(horizontal="center")

        ws.merge_cells("A2:F2")
        ws["A2"] = (f"Active ballots: {rd['total_active']}   |   "
                    f"Majority needed: {rd['majority_needed']}   |   "
                    f"15% threshold: {rd['total_active'] * 0.15:.1f} votes   |   "
                    f"Exhausted: {rd.get('exhausted', 0)}")
        ws["A2"].font      = Font(name="Calibri", italic=True, size=10, color="595959")
        ws["A2"].alignment = Alignment(horizontal="center")

        # Row 3: elimination rule in effect for this round
        elim_rule = rd.get("elim_reason", "")
        if not elim_rule:
            seat_rn_val = rd.get("seat_round_num", 1)
            if seat_rn_val == 1:
                elim_rule = "15% elimination threshold"
            elif isinstance(seat_rn_val, str):
                elim_rule = f"{seat_rn_val} runoff ballot"
            else:
                elim_rule = "Lowest vote-getter eliminated"
        ws.merge_cells("A3:F3")
        ws["A3"] = f"Elimination rule: {elim_rule}"
        ws["A3"].font      = Font(name="Calibri", italic=True, size=10, color="7F5A00")
        ws["A3"].alignment = Alignment(horizontal="center")

        headers = ["#", "Candidate", "Votes", "% of Active", "Majority?", "Status"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=5, column=col, value=h)
            sc(cell, fill=hdr_fill, font=hdr_font, align="center")

        sorted_cands = sorted(rd["vote_counts"].items(), key=lambda x: -x[1])
        for row_i, (num, votes) in enumerate(sorted_cands, 6):
            name = candidate_map.get(num, f"Candidate {num}")
            pct  = votes / rd["total_active"] * 100 if rd["total_active"] > 0 else 0
            maj  = "YES ✓" if votes >= rd["majority_needed"] else "No"

            if num in rd["elected"]:
                status = "ELECTED"
                fill   = elect_fill
                font   = bold_font
            elif num in rd["eliminated"]:
                status = f"ELIMINATED — {rd.get('elim_reason','')}"
                fill   = elim_fill
                font   = norm_font
            else:
                status = "Continuing"
                fill   = None
                font   = norm_font

            vals = [num, name, votes, f"{pct:.1f}%", maj, status]
            for col, val in enumerate(vals, 1):
                cell = ws.cell(row=row_i, column=col, value=val)
                sc(cell, fill=fill, font=font, align="center" if col != 2 else "left")

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 30
        for ch in ["C","D","E","F"]:
            ws.column_dimensions[ch].width = 16

    # ── Summary sheet ────────────────────────────────────────────────────────
    ws_s = wb.create_sheet(title="Summary", index=0)
    ws_s.merge_cells("A1:E1")
    ws_s["A1"]           = f"{election_name} — Final Results"
    ws_s["A1"].font      = Font(name="Calibri", bold=True, size=14, color="1F3864")
    ws_s["A1"].alignment = Alignment(horizontal="center")

    ws_s.merge_cells("A2:E2")
    ws_s["A2"] = (f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}   |   "
                  f"{DISTRICT}")
    ws_s["A2"].font      = Font(name="Calibri", italic=True, size=10, color="595959")
    ws_s["A2"].alignment = Alignment(horizontal="center")

    for col, h in enumerate(["#", "Candidate", "Seat", "Elected on Distribution Round", "Votes (When Elected)"], 1):
        cell = ws_s.cell(row=4, column=col, value=h)
        sc(cell, fill=hdr_fill, font=hdr_font, align="center")

    elected_summary = []
    for rd in rounds_data:
        for num in rd["elected"]:
            votes    = rd["vote_counts"].get(num, 0)
            seat_n   = rd.get("seat_num", "")
            seat_rn  = rd.get("seat_round_num", "")
            elected_summary.append((num, seat_n, seat_rn, votes))

    for row_i, (num, seat, seat_rnd, votes) in enumerate(elected_summary, 5):
        name = candidate_map.get(num, f"Candidate {num}")
        vals = [row_i - 4, name, f"Seat {seat}" if seat else "", seat_rnd, votes]
        for col, val in enumerate(vals, 1):
            cell = ws_s.cell(row=row_i, column=col, value=val)
            sc(cell, fill=elect_fill, align="center" if col != 2 else "left")

    ws_s.column_dimensions["A"].width = 5
    ws_s.column_dimensions["B"].width = 30
    for ch in ["C","D","E"]:
        ws_s.column_dimensions[ch].width = 20

    wb.save(output_path)
    print(c(f"\n  Excel report saved → {output_path}", Color.GREEN))
    log(f"[EXCEL] Saved: {output_path}", also_print=False)

# =============================================================================
# OUTPUT DIRECTORY
# =============================================================================

def get_output_dir(script_dir):
    default = os.path.join(script_dir, "election_results")
    print(c(f"\n  Results will be saved to:", Color.BOLD))
    print(c(f"    {default}", Color.CYAN))
    if not confirm("  Use this location?"):
        custom = input("  Enter full folder path: ").strip().strip('"').strip("'")
        if custom:
            default = custom
    os.makedirs(default, exist_ok=True)
    return default

# =============================================================================
# MAIN
# =============================================================================

def main():
    banner()

    log(f"SESSION STARTED: {datetime.datetime.now().isoformat()}", also_print=False)
    log(f"CONVENTION: {DISTRICT} — {CONVENTION_DATE}", also_print=False)
    log(f"TABULATOR VERSION: {VERSION}", also_print=False)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = get_output_dir(script_dir)
    log(f"OUTPUT DIR: {output_dir}", also_print=False)

    all_results = {}

    while True:
        section_header("MAIN MENU")
        choice = menu("Select an election to run:", [
            "A.  State Central Committee — Women           (4 seats, IRV)",
            "B.  State Central Committee — Men/Non-Binary  (4 seats, IRV)",
            "C.  DEI Committee Chair                        (1 seat, IRV)",
            "D.  State Convention Committee                 (14 seats, slate vote)",
            "Save audit log and exit",
        ])

        if choice == 1:
            result = load_ballots_for_election("SCC — Women", script_dir)
            if result[0] is not None:
                ballots, cmap, spoiled = result
                elected, _, _ = run_scc_election(
                    "Women", 4, ballots, cmap, spoiled, output_dir)
                all_results["SCC_Women"] = [cmap.get(n, n) for n in elected]

        elif choice == 2:
            result = load_ballots_for_election("SCC — Men/Non-Binary", script_dir)
            if result[0] is not None:
                ballots, cmap, spoiled = result
                elected, _, _ = run_scc_election(
                    "Men_NonBinary", 4, ballots, cmap, spoiled, output_dir)
                all_results["SCC_MNB"] = [cmap.get(n, n) for n in elected]

        elif choice == 3:
            result = load_ballots_for_election("DEI Committee Chair", script_dir)
            if result[0] is not None:
                ballots, cmap, spoiled = result
                winner = run_dei_election(ballots, cmap, spoiled, output_dir)
                all_results["DEI_Chair"] = cmap.get(winner, winner) if winner else None

        elif choice == 4:
            slate_approved = run_committee_election(output_dir)
            all_results["Convention_Committee_Slate"] = "Approved" if slate_approved else "Failed"

        elif choice == 5:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            log_path  = os.path.join(output_dir, f"audit_log_{timestamp}.txt")
            save_audit_log(log_path)
            print(c("\n  ══ SESSION COMPLETE ══", Color.BOLD + Color.CYAN))
            print(c("\n  Elections completed this session:", Color.WHITE))
            for k, v in all_results.items():
                print(c(f"    {k}: {v}", Color.GREEN))
            print(c(f"\n  All files saved to: {output_dir}", Color.CYAN))
            break


if __name__ == "__main__":
    main()
