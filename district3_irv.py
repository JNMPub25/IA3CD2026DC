#!/usr/bin/env python3
"""
=============================================================================
district3_irv.py — Instant Runoff Voting (IRV) / Ranked Choice Module
=============================================================================
Core ranked-choice voting logic for the 3rd Congressional District
Democratic Convention tabulation system.

Functions:
  - count_round_votes()            — tally first-active-choice votes
  - compute_threshold_elimination() — 15% threshold (Step A)
  - drop_lowest_candidate()         — drop lowest vote-getter (Step B)
  - run_irv_one_seat()              — full IRV cycle for a single seat

Two-step Round 1 process:
  Step A: Apply 15% threshold → redistribute → check for 50%+1 winner
  Step B: If no winner, drop the lowest vote-getter → redistribute

Round 2+: Only the lowest vote-getter is dropped (no threshold).

Depends on district3_runoff.run_runoff() for tie resolution.
=============================================================================
"""

from district3_runoff import run_runoff


# =============================================================================
# CONSTANTS  (duplicated here to keep module self-contained)
# =============================================================================

ELIMINATION_THRESHOLD = 0.15   # 15% — candidates below this are eliminated
MAJORITY_THRESHOLD    = 0.50   # must receive MORE THAN 50% to be elected


# =============================================================================
# DISPLAY / LOGGING HELPERS  (injected by the main tabulator at init)
# =============================================================================

_c    = lambda text, color: text
_log  = lambda text, **kw: None
_seat_ballot_header       = lambda seat, label: None
_distribution_round_header = lambda rnd, seat, label: None
_display_round_table      = lambda *a, **kw: None
_Color = None


def init_irv_display(c_func, log_func, seat_ballot_header_func,
                     distribution_round_header_func,
                     display_round_table_func, color_class):
    """
    Called once by the main tabulator to inject display/logging functions.
    """
    global _c, _log, _seat_ballot_header, _distribution_round_header
    global _display_round_table, _Color
    _c                         = c_func
    _log                       = log_func
    _seat_ballot_header        = seat_ballot_header_func
    _distribution_round_header = distribution_round_header_func
    _display_round_table       = display_round_table_func
    _Color                     = color_class


# =============================================================================
# VOTE COUNTING
# =============================================================================

def count_round_votes(ballots, active_nums):
    """
    Count first-active-choice votes for one IRV round.
    Returns:
      vote_counts   : {cand_num_str: int}   (only active candidates)
      total_active  : int  (ballots with an active choice)
      exhausted     : int  (ballots with no active choice remaining)
    """
    vote_counts = {num: 0 for num in active_nums}
    exhausted   = 0

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


# =============================================================================
# ELIMINATION FUNCTIONS  (two-step approach)
# =============================================================================

def compute_threshold_elimination(vote_counts,
                                  threshold=ELIMINATION_THRESHOLD):
    """
    STEP A — Apply the 15% threshold rule.

    Eliminates candidates with < 15% of total active votes, subject to the
    safeguard that no more than 49% of remaining candidates may be eliminated.

    This function does NOT drop the lowest vote-getter.  After calling this,
    the caller should redistribute votes and check for a 50%+1 winner before
    proceeding to drop_lowest_candidate() as a separate step.

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

    # Step 2: Safeguard — cannot eliminate >= 50% of candidates (III-C-4)
    max_eliminate = max(1, (n - 1) // 2)

    if len(below_threshold) > max_eliminate:
        # Take the bottom max_eliminate candidates by vote count
        tentative = [num for num, v in sorted_cands[:max_eliminate]]
        # Check if the cut line splits a tied group — if the last
        # candidate to be eliminated has the same votes as the first
        # candidate to be kept, pull back to avoid splitting the tie
        if tentative:
            boundary_votes = vote_counts[tentative[-1]]
            # Remove all candidates at the boundary vote count
            # (they stay in the race with their tied peers)
            below_threshold = [num for num in tentative
                               if vote_counts[num] < boundary_votes]
            if below_threshold:
                cut_votes  = vote_counts[below_threshold[-1]]
                actual_pct = cut_votes / total * 100
            # If pulling back leaves nobody to eliminate, that's OK —
            # Step B (drop lowest) will handle the next elimination
        else:
            below_threshold = []

    return below_threshold, actual_pct


def drop_lowest_candidate(vote_counts):
    """
    STEP B — Identify the lowest vote-getter(s).

    Called only when the threshold step did not produce a winner.
    Returns a list of candidate number strings tied for lowest.
    If there is a tie for lowest, returns ALL tied candidates (the caller
    should invoke a tiebreak runoff if eliminating all would be problematic).
    """
    if not vote_counts:
        return []
    sorted_cands = sorted(vote_counts.items(), key=lambda x: x[1])
    lowest_votes = sorted_cands[0][1]
    tied_lowest  = [num for num, v in vote_counts.items()
                    if v == lowest_votes]
    return tied_lowest


# =============================================================================
# SINGLE-SEAT IRV
# =============================================================================

def run_irv_one_seat(ballots, candidate_map, active_nums, election_label,
                     seat_num, max_wins_this_seat, rounds_data_out,
                     election_key=""):
    """
    Run IRV for a single seat among the given active_nums.

    Two-step Round 1 process:
      Step A: 15% threshold elimination → redistribute → winner check
      Step B: Drop lowest vote-getter (only if Step A didn't produce a winner)

    Round 2+: Only the lowest vote-getter is dropped.

    Appends round data dicts to rounds_data_out.
    Returns winner_num (str) or None.
    """
    active     = list(active_nums)
    seat_round = 1

    _seat_ballot_header(seat_num, election_label)

    while active:
        _distribution_round_header(seat_round, seat_num, election_label)

        vote_counts, total_active, exhausted = count_round_votes(
            ballots, active)
        majority_needed = int(total_active * MAJORITY_THRESHOLD) + 1

        # Check for winner (> 50%)
        potential_winners = [num for num, v in vote_counts.items()
                             if v >= majority_needed]

        # Cap winners for the first seat's first round
        if seat_round == 1 and seat_num == 1:
            potential_winners = potential_winners[:max_wins_this_seat]

        winner_num = potential_winners[0] if potential_winners else None

        if winner_num:
            _display_round_table(candidate_map, vote_counts, total_active,
                                 majority_needed, elected=[winner_num])
            rounds_data_out.append({
                "seat_num":        seat_num,
                "seat_round_num":  seat_round,
                "vote_counts":     dict(vote_counts),
                "total_active":    total_active,
                "exhausted":       exhausted,
                "majority_needed": majority_needed,
                "elected":         [winner_num],
                "eliminated":      [],
            })
            name = candidate_map.get(winner_num, winner_num)
            print(_c(f"\n  ✓ {name} ELECTED (Seat {seat_num})",
                     _Color.GREEN + _Color.BOLD))
            _log(f"[SEAT {seat_num} — Round {seat_round}] WINNER: {name}",
                 also_print=False)
            return winner_num

        # ── No winner — determine eliminations ─────────────────────────

        if seat_round == 1:
            # ── STEP A: 15% threshold ──────────────────────────────────
            threshold_elim, eff_pct = compute_threshold_elimination(
                vote_counts)

            if threshold_elim:
                elim_reason = (f"below {eff_pct:.1f}% "
                               f"elimination threshold")
                _display_round_table(
                    candidate_map, vote_counts, total_active,
                    majority_needed, eliminated=threshold_elim)
                rounds_data_out.append({
                    "seat_num":        seat_num,
                    "seat_round_num":  seat_round,
                    "vote_counts":     dict(vote_counts),
                    "total_active":    total_active,
                    "exhausted":       exhausted,
                    "majority_needed": majority_needed,
                    "elected":         [],
                    "eliminated":      list(threshold_elim),
                    "elim_reason":     elim_reason,
                })
                elim_names = [candidate_map.get(n, n)
                              for n in threshold_elim]
                _log(f"[SEAT {seat_num} — Round {seat_round} Step A] "
                     f"Threshold eliminated: {elim_names} ({elim_reason})",
                     also_print=False)
                for num in threshold_elim:
                    name = candidate_map.get(num, num)
                    active.remove(num)
                    print(_c(f"  ✗ {name} eliminated  ({elim_reason})",
                             _Color.RED))

                if len(active) == 1:
                    winner_num = active[0]
                    name = candidate_map.get(winner_num, winner_num)
                    print(_c(f"\n  Only one candidate remaining "
                             f"— {name} is elected.", _Color.YELLOW))
                    print(_c(f"  ✓ {name} ELECTED (Seat {seat_num})",
                             _Color.GREEN + _Color.BOLD))
                    _log(f"[SEAT {seat_num}] Last candidate standing: "
                         f"{name}", also_print=False)
                    # Record the final round
                    vc2, ta2, ex2 = count_round_votes(ballots, active)
                    rounds_data_out.append({
                        "seat_num":       seat_num,
                        "seat_round_num": seat_round + 1,
                        "vote_counts":    dict(vc2),
                        "total_active":   ta2,
                        "exhausted":      ex2,
                        "majority_needed": int(ta2 * MAJORITY_THRESHOLD) + 1,
                        "elected":        [winner_num],
                        "eliminated":     [],
                    })
                    return winner_num

                # Recount after threshold redistribution
                vote_counts, total_active, exhausted = count_round_votes(
                    ballots, active)
                majority_needed = int(
                    total_active * MAJORITY_THRESHOLD) + 1

                # Check for winner after threshold redistribution
                potential_winners = [
                    num for num, v in vote_counts.items()
                    if v >= majority_needed]
                if potential_winners:
                    winner_num = potential_winners[0]
                    _display_round_table(
                        candidate_map, vote_counts, total_active,
                        majority_needed, elected=[winner_num])
                    rounds_data_out.append({
                        "seat_num":        seat_num,
                        "seat_round_num":  seat_round,
                        "vote_counts":     dict(vote_counts),
                        "total_active":    total_active,
                        "exhausted":       exhausted,
                        "majority_needed": majority_needed,
                        "elected":         [winner_num],
                        "eliminated":      [],
                    })
                    name = candidate_map.get(winner_num, winner_num)
                    print(_c(f"\n  ✓ {name} ELECTED (Seat {seat_num}) "
                             f"— winner after threshold redistribution",
                             _Color.GREEN + _Color.BOLD))
                    _log(f"[SEAT {seat_num} — Round {seat_round} Step B] "
                         f"WINNER after threshold: {name}",
                         also_print=False)
                    return winner_num

            # ── STEP B: Drop the lowest vote-getter ────────────────────
            tied_lowest = drop_lowest_candidate(vote_counts)

            if len(tied_lowest) == 1:
                to_eliminate = tied_lowest
                lowest_votes = vote_counts[tied_lowest[0]]
                lowest_pct = (lowest_votes / total_active * 100
                              if total_active > 0 else 0)
                elim_reason = (f"lowest vote-getter ({lowest_pct:.1f}%) "
                               f"— Round 1 drop")
            else:
                # Tie for lowest — runoff ballot required
                tied_names = [candidate_map.get(n, n) for n in tied_lowest]
                print(_c(f"\n  Tie for lowest ({', '.join(tied_names)}) "
                         f"— runoff ballot required.",
                         _Color.YELLOW + _Color.BOLD))
                _log(f"[SEAT {seat_num} — Round {seat_round}] "
                     f"TIE for lowest: {tied_names}", also_print=False)

                _display_round_table(
                    candidate_map, vote_counts, total_active,
                    majority_needed, eliminated=tied_lowest)
                rounds_data_out.append({
                    "seat_num":        seat_num,
                    "seat_round_num":  seat_round,
                    "vote_counts":     dict(vote_counts),
                    "total_active":    total_active,
                    "exhausted":       exhausted,
                    "majority_needed": majority_needed,
                    "elected":         [],
                    "eliminated":      list(tied_lowest),
                    "elim_reason":     ("tie for lowest — "
                                       "tiebreaker runoff ballot issued"),
                })

                tied_map = {n: candidate_map[n] for n in tied_lowest}
                tb_winner = run_runoff(
                    "tiebreaker", tied_map, election_key, election_label,
                    seat_num, "Tiebreaker", rounds_data_out)
                if tb_winner is None:
                    return None

                to_elim_tb = [n for n in tied_lowest if n != tb_winner]
                for num in to_elim_tb:
                    active.remove(num)
                    print(_c(f"  ✗ {candidate_map.get(num, num)} eliminated"
                             f"  (tiebreaker runoff)", _Color.RED))
                    _log(f"[SEAT {seat_num}] Tiebreaker loser: "
                         f"{candidate_map.get(num, num)}",
                         also_print=False)

                if len(active) == 2:
                    h2h_map = {n: candidate_map[n] for n in active}
                    h2h_winner = run_runoff(
                        "head-to-head", h2h_map, election_key,
                        election_label, seat_num, "H2H",
                        rounds_data_out)
                    if h2h_winner:
                        return h2h_winner

                seat_round += 1
                continue

        else:
            # ── ROUND 2+: drop the single lowest ──────────────────────
            sorted_by_votes = sorted(vote_counts.items(),
                                     key=lambda x: x[1])
            lowest_votes    = sorted_by_votes[0][1]
            tied_for_lowest = [num for num, v in vote_counts.items()
                               if v == lowest_votes]

            if len(tied_for_lowest) == 1:
                to_eliminate = tied_for_lowest
                lowest_pct = (lowest_votes / total_active * 100
                              if total_active > 0 else 0)
                elim_reason = f"lowest vote-getter ({lowest_pct:.1f}%)"

            else:
                # Tie for lowest — runoff ballot required
                tied_names = [candidate_map.get(n, n)
                              for n in tied_for_lowest]
                print(_c(f"\n  Tie for lowest ({', '.join(tied_names)}) "
                         f"— runoff ballot required.",
                         _Color.YELLOW + _Color.BOLD))
                _log(f"[SEAT {seat_num} — Round {seat_round}] "
                     f"TIE for lowest: {tied_names}", also_print=False)

                _display_round_table(
                    candidate_map, vote_counts, total_active,
                    majority_needed, eliminated=tied_for_lowest)
                rounds_data_out.append({
                    "seat_num":        seat_num,
                    "seat_round_num":  seat_round,
                    "vote_counts":     dict(vote_counts),
                    "total_active":    total_active,
                    "exhausted":       exhausted,
                    "majority_needed": majority_needed,
                    "elected":         [],
                    "eliminated":      list(tied_for_lowest),
                    "elim_reason":     ("tie for lowest — "
                                       "tiebreaker runoff ballot issued"),
                })

                tied_map = {n: candidate_map[n] for n in tied_for_lowest}
                tb_winner = run_runoff(
                    "tiebreaker", tied_map, election_key, election_label,
                    seat_num, "Tiebreaker", rounds_data_out)
                if tb_winner is None:
                    return None

                to_elim_tb = [n for n in tied_for_lowest
                              if n != tb_winner]
                for num in to_elim_tb:
                    active.remove(num)
                    print(_c(f"  ✗ {candidate_map.get(num, num)} eliminated"
                             f"  (tiebreaker runoff)", _Color.RED))
                    _log(f"[SEAT {seat_num}] Tiebreaker loser: "
                         f"{candidate_map.get(num, num)}",
                         also_print=False)

                if len(active) == 2:
                    h2h_map = {n: candidate_map[n] for n in active}
                    h2h_winner = run_runoff(
                        "head-to-head", h2h_map, election_key,
                        election_label, seat_num, "H2H",
                        rounds_data_out)
                    if h2h_winner:
                        return h2h_winner

                seat_round += 1
                continue

        # ── Normal elimination (clear lowest, no tie) ──────────────────
        _display_round_table(candidate_map, vote_counts, total_active,
                             majority_needed, eliminated=to_eliminate)
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
        _log(f"[SEAT {seat_num} — Round {seat_round}] No winner.  "
             f"Eliminated: {elim_names} ({elim_reason})",
             also_print=False)

        for num in to_eliminate:
            name = candidate_map.get(num, num)
            active.remove(num)
            print(_c(f"  ✗ {name} eliminated  ({elim_reason})",
                     _Color.RED))

        if len(active) == 1:
            winner_num = active[0]
            name = candidate_map.get(winner_num, winner_num)
            print(_c(f"\n  Only one candidate remaining "
                     f"— {name} is elected.", _Color.YELLOW))
            print(_c(f"  ✓ {name} ELECTED (Seat {seat_num})",
                     _Color.GREEN + _Color.BOLD))
            _log(f"[SEAT {seat_num}] Last candidate standing: {name}",
                 also_print=False)
            vc_final, ta_final, ex_final = count_round_votes(
                ballots, active)
            rounds_data_out.append({
                "seat_num":       seat_num,
                "seat_round_num": seat_round + 1,
                "vote_counts":    dict(vc_final),
                "total_active":   ta_final,
                "exhausted":      ex_final,
                "majority_needed": int(ta_final * MAJORITY_THRESHOLD) + 1,
                "elected":        [winner_num],
                "eliminated":     [],
            })
            return winner_num

        seat_round += 1

    return None
