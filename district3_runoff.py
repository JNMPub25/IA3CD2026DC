#!/usr/bin/env python3
"""
=============================================================================
district3_runoff.py — Runoff / Simple Ballot Module
=============================================================================
Handles tiebreaker and head-to-head runoff ballots for the 3rd Congressional
District Democratic Convention tabulation system.

Triggered when two or more candidates share the lowest vote total and one
must be dropped, or when only 2 candidates remain and a head-to-head
ballot is called.

Both ballot types use ballot.html with ?type=runoff in the URL.
Results arrive as a separate Google Sheets CSV export (Runoff tab).
=============================================================================
"""

import csv


# =============================================================================
# MODULE-LEVEL STATE
# =============================================================================

_ballot_base_url = ""   # set once per session when first runoff is needed


# =============================================================================
# DISPLAY / LOGGING HELPERS  (injected by the main tabulator at init)
# =============================================================================
# These are set by init_runoff_display() so this module can print formatted
# output without importing the main tabulator (avoids circular imports).

_c    = lambda text, color: text          # colorize function
_log  = lambda text, **kw: None           # audit log function
_section_header = lambda title: None      # section header display
_Color = None                              # Color class reference
_get_csv_filepath = lambda label: None    # CSV file path prompt
_confirm = lambda msg: True               # yes/no prompt


def init_runoff_display(c_func, log_func, section_header_func,
                        color_class, get_csv_func, confirm_func):
    """
    Called once by the main tabulator to inject display/logging functions.
    This avoids circular imports between modules.
    """
    global _c, _log, _section_header, _Color, _get_csv_filepath, _confirm
    _c                = c_func
    _log              = log_func
    _section_header   = section_header_func
    _Color            = color_class
    _get_csv_filepath = get_csv_func
    _confirm          = confirm_func


# =============================================================================
# RUNOFF URL GENERATION
# =============================================================================

def get_ballot_base_url():
    """Prompt operator for the ballot.html URL (cached for session)."""
    global _ballot_base_url
    if not _ballot_base_url:
        print(_c("\n  A runoff ballot needs to be generated.", _Color.CYAN))
        print(_c("  Enter the GitHub Pages URL where ballot.html is hosted,",
                 _Color.GREY))
        print(_c("  e.g.  https://yourname.github.io/ballot.html", _Color.GREY))
        url = input("\n  Ballot URL: ").strip().rstrip("/")
        _ballot_base_url = url
    return _ballot_base_url


def make_runoff_url(base_url, runoff_election_key, candidate_map_subset):
    """Build a ballot.html URL for a two-candidate runoff (type=runoff)."""
    cands_param = "-".join(
        f"{num}-{name.replace(' ', '%20')}"
        for num, name in candidate_map_subset.items()
    )
    return (f"{base_url}?election={runoff_election_key}"
            f"&candidates={cands_param}&seats=1&type=runoff")


def show_runoff_qr(url, label):
    """Print the runoff URL and show a QR code if the qrcode library is available."""
    print(_c(f"\n  ┌─ {label} ─────────────────────────────────────────────",
             _Color.MAGENTA))
    print(_c(f"  │  URL: {url}", _Color.CYAN))
    print(_c(f"  └────────────────────────────────────────────────────────",
             _Color.MAGENTA))
    _log(f"[RUNOFF URL — {label}]: {url}", also_print=False)
    try:
        import qrcode as _qr
        qr = _qr.QRCode(box_size=2, border=2)
        qr.add_data(url)
        qr.make(fit=True)
        qr.print_ascii(invert=True)
    except ImportError:
        print(_c("  (Install qrcode[pil] to display QR code here, or use "
                 "district3_setup.py to generate one.)", _Color.GREY))


# =============================================================================
# RUNOFF CSV LOADING
# =============================================================================

def load_runoff_csv(filepath):
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
                (fn for fn in fieldnames if fn.strip().lower() == "choice"),
                None)
            if not choice_col:
                warnings.append(
                    "No 'Choice' column found in the runoff CSV. "
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


# =============================================================================
# RUNOFF ORCHESTRATOR
# =============================================================================

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

    _section_header(f"RUNOFF — {election_label} — Seat {seat_num} — "
                    f"{runoff_type.replace('-',' ').title()}")

    if runoff_type == "tiebreaker":
        print(_c(f"\n  Tie for lowest between: {names[0]}  and  {names[1]}",
                 _Color.MAGENTA + _Color.BOLD))
        print(_c("  A tiebreaker runoff ballot will be issued. "
                 "Delegates select ONE candidate.", _Color.GREY))
    else:
        print(_c(f"\n  Head-to-head ballot: {names[0]}  vs  {names[1]}",
                 _Color.MAGENTA + _Color.BOLD))
        print(_c("  Delegates select ONE candidate. "
                 "The winner is elected to this seat.", _Color.GREY))

    # Generate and display URL + QR code
    runoff_key = f"runoff-{election_key}-{''.join(nums)}"
    base_url   = get_ballot_base_url()
    url        = make_runoff_url(base_url, runoff_key, candidate_map_subset)
    show_runoff_qr(url, f"Seat {seat_num} — {runoff_type.title()}")

    print(_c("\n  ─── Chair instructions ───────────────────────────────────────",
             _Color.GREY))
    print(_c("  1. Project the QR code or read the URL aloud.", _Color.GREY))
    print(_c("  2. Announce to delegates:", _Color.WHITE))
    print(_c(f'     "Which candidate do you approve — {names[0]} or {names[1]}?'
             f' You may select only one."', _Color.BOLD))
    print(_c("  3. Allow time to vote, then close the ballot window.",
             _Color.GREY))
    print(_c("  4. Export the Runoff tab from Google Sheets as CSV.",
             _Color.GREY))
    print(_c("  ─────────────────────────────────────────────────────────────",
             _Color.GREY))

    input(_c("\n  Press Enter when ready to load the runoff CSV...",
             _Color.YELLOW))
    path = _get_csv_filepath(f"Runoff CSV — {runoff_type}")

    vote_counts = {}
    total       = 0
    winner_num  = None

    if path:
        vote_counts, warnings = load_runoff_csv(path)
        for w in warnings:
            print(_c(f"  ⚠  {w}", _Color.YELLOW))

        if vote_counts:
            total = sum(vote_counts.values())
            print(_c(f"\n  Results ({total} votes cast):", _Color.BOLD))
            for num in sorted(nums, key=lambda n: -vote_counts.get(n, 0)):
                name = candidate_map_subset[num]
                cnt  = vote_counts.get(num, 0)
                pct  = cnt / total * 100 if total else 0
                bar  = "█" * int(pct / 5)
                print(_c(f"    {name:<28} {cnt:>4}  ({pct:>5.1f}%)  {bar}",
                         _Color.WHITE))
                _log(f"  [RUNOFF {runoff_type.upper()}] {name}: {cnt} ({pct:.1f}%)",
                     also_print=False)

            max_votes      = max(vote_counts.values())
            tied_in_runoff = [n for n, v in vote_counts.items()
                              if v == max_votes]

            if len(tied_in_runoff) == 1:
                winner_num = tied_in_runoff[0]
            else:
                print(_c("\n  ⚠  Runoff itself is tied. "
                         "Chair must determine the winner.", _Color.RED))
                _log("[RUNOFF TIE] Runoff result tied — "
                     "Chair determination required.", also_print=False)
                print(_c("  Enter the winner:", _Color.YELLOW))
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
                    print(_c("  Please enter a valid number.", _Color.RED))
    else:
        # No CSV — manual entry fallback
        print(_c("\n  No CSV loaded. Enter the winner manually:",
                 _Color.YELLOW))
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
            print(_c("  Please enter a valid number.", _Color.RED))
        vote_counts = {n: 0 for n in nums}

    if winner_num:
        loser_nums  = [n for n in nums if n != winner_num]
        winner_name = candidate_map_subset[winner_num]
        loser_names = [candidate_map_subset[n] for n in loser_nums]

        if runoff_type == "head-to-head":
            print(_c(f"\n  ✓ {winner_name} ELECTED "
                     f"(Seat {seat_num} — head-to-head runoff)",
                     _Color.GREEN + _Color.BOLD))
        else:
            print(_c(f"\n  {winner_name} advances.  "
                     f"{', '.join(loser_names)} eliminated.", _Color.CYAN))

        _log(f"[RUNOFF {runoff_type.upper()} — Seat {seat_num}] "
             f"Winner: {winner_name}  Eliminated: {loser_names}",
             also_print=False)

        rounds_data_out.append({
            "seat_num":        seat_num,
            "seat_round_num":  seat_round_label,
            "vote_counts":     {n: vote_counts.get(n, 0) for n in nums},
            "total_active":    total,
            "exhausted":       0,
            "majority_needed": total // 2 + 1 if total else 1,
            "elected":         [winner_num]
                               if runoff_type == "head-to-head" else [],
            "eliminated":      loser_nums,
            "elim_reason":     f"{runoff_type} runoff ballot",
        })

    return winner_num
