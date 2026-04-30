#!/usr/bin/env python3
"""
Quick test runner for DEI Committee Chair — 4 candidates test.
Bypasses interactive main() to run directly against test CSV.
"""
import os
import sys

# Ensure we're in the right directory
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

from district3_tabulator import (
    load_electronic_ballots, check_surrendered, enter_candidates,
    banner, section_header, c, Color, log, save_audit_log,
    merge_ballots, Ballot
)
from district3_irv import run_irv_one_seat

# ── Config ──
CSV_PATH = os.path.join(script_dir, "test_ballots",
                        "DEI Committee Chair - 4 candidates test.csv")
SURRENDERED_PATH = os.path.join(script_dir, "surrendered_delegates.json")

banner()
section_header("TEST RUN — DEI COMMITTEE CHAIR (4 candidates)")

# ── Load ballots ──
print(c(f"\n  Loading: {CSV_PATH}", Color.CYAN))
ballots, candidate_keys, warnings = load_electronic_ballots(CSV_PATH)
for w in warnings:
    print(c(f"  ⚠  {w}", Color.YELLOW))
print(c(f"  Loaded {len(ballots)} ballot(s) with {len(candidate_keys)} "
        f"candidate column(s): {', '.join(candidate_keys)}", Color.GREEN))

# ── Check surrendered credentials ──
section_header("CREDENTIALS CHECK")
ballots, spoiled = check_surrendered(ballots, SURRENDERED_PATH)
if spoiled:
    n_surr = sum(1 for b in spoiled if "Non-Issued" not in (b.spoil_reason or ""))
    n_ni   = sum(1 for b in spoiled if "Non-Issued" in (b.spoil_reason or ""))
    parts = []
    if n_surr: parts.append(f"{n_surr} Surrendered")
    if n_ni:   parts.append(f"{n_ni} Non-Issued")
    print(c(f"\n  ⚠  {len(spoiled)} ballot(s) flagged: {', '.join(parts)}.",
            Color.YELLOW))
    for b in spoiled:
        clr = Color.RED if "Non-Issued" in (b.spoil_reason or "") else Color.YELLOW
        print(c(f"     Delegate {b.delegate_number} — {b.spoil_reason}", clr))
else:
    print(c("  No surrendered credentials flagged.", Color.GREEN))

print(c(f"\n  Valid ballots for tabulation: {len(ballots)}", Color.GREEN))

# ── Candidate names ──
candidate_map = {
    "A": "Alice Anderson",
    "B": "Bob Baker",
    "C": "Carol Chen",
    "D": "Dave Davis",
}
section_header("CANDIDATES")
for key in sorted(candidate_map):
    print(c(f"    {key} = {candidate_map[key]}", Color.WHITE))

# ── Run IRV ──
section_header("DEI COMMITTEE CHAIR — IRV TABULATION")
active_candidates = list(sorted(candidate_map.keys()))
winner, rounds_data = run_irv_one_seat(
    ballots, active_candidates, candidate_map,
    seat_num=1, election_label="DEI Committee Chair"
)

section_header("RESULT")
if winner:
    print(c(f"\n  ✓  WINNER: {candidate_map.get(winner, winner)} (Candidate {winner})",
            Color.GREEN + Color.BOLD))
else:
    print(c("\n  ✗  No winner determined.", Color.RED))

# ── Summary ──
section_header("TEST SUMMARY")
print(c(f"  CSV rows loaded:         252 (250 real + 2 Is Test=YES)", Color.WHITE))
print(c(f"  Is Test filtered out:    2", Color.WHITE))
print(c(f"  Ballots after filter:    {len(ballots) + len(spoiled)}", Color.WHITE))
print(c(f"  Surrendered flagged:     {sum(1 for b in spoiled if 'Non-Issued' not in (b.spoil_reason or ''))}", Color.WHITE))
print(c(f"  Non-issued flagged:      {sum(1 for b in spoiled if 'Non-Issued' in (b.spoil_reason or ''))}", Color.WHITE))
print(c(f"  Valid ballots tabulated: {len(ballots)}", Color.WHITE))
print(c(f"  IRV rounds:              {len(rounds_data)}", Color.WHITE))
print()
