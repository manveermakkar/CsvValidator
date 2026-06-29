"""
core/fuzzy_aligner.py  —  Phase 2 (NEW)
----------------------------------------
Fuzzy-logic row alignment powered by rapidfuzz.

Performance strategy
--------------------
1. **Exact match first** (O(n) hash lookup) — the vast majority of rows are
   matched instantly without any fuzzy computation.
2. **Candidate pool reduction** — only the *unmatched* rows from both files
   enter the fuzzy stage, keeping the search space small.
3. **rapidfuzz** — uses Cython-compiled C++ under the hood; 10–100× faster
   than pure-Python thefuzz/difflib for large string sets.
4. **Blocking by first character** — when the unmatched pool is very large
   (> BLOCK_THRESHOLD), candidates are pre-filtered to those sharing the
   same first character as the query key before calling extractOne.
5. **Score threshold gate** — matches scoring below `threshold` (default 80)
   are treated as unmatched rather than force-paired with a bad candidate.
6. **Each B row used at most once** — the matched B candidate is removed from
   the pool after assignment, preserving 1-to-1 row correspondence.

Returns
-------
aligned_a, aligned_b : pd.DataFrame  — aligned with STATUS_COL injected
match_log            : list[dict]     — one entry per row:
    {key_a, key_b, score, match_type}
    match_type ∈ {'exact', 'fuzzy', 'unmatched_a', 'unmatched_b'}
"""

import pandas as pd
import numpy as np
from rapidfuzz import process, fuzz

from core.aligner import STATUS_COL

_SEP             = "|||"   # separator for composite key concatenation
BLOCK_THRESHOLD  = 500     # pool size above which first-char blocking is applied


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def align_fuzzy(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    key_cols: list[str],
    threshold: int = 80,
) -> tuple[pd.DataFrame, pd.DataFrame, list[dict]]:
    """
    Fuzzy-match rows between df_a and df_b using *key_cols* as identifiers.

    Parameters
    ----------
    df_a, df_b : pd.DataFrame
        Already column-mapped DataFrames (same column set, string dtype).
    key_cols   : list[str]
        One or more column names used as the matching key.
        Multiple columns are concatenated with '|||'.
    threshold  : int (0–100)
        Minimum rapidfuzz score to accept a fuzzy match.  Rows below this
        score are flagged as unmatched rather than force-paired.

    Returns
    -------
    aligned_a : pd.DataFrame
    aligned_b : pd.DataFrame
    match_log : list[dict]
    """
    # ── Build composite key strings ─────────────────────────────────────────
    keys_a = _make_key(df_a, key_cols).reset_index(drop=True)
    keys_b = _make_key(df_b, key_cols).reset_index(drop=True)

    # ── Phase 1: Exact matching (O(n) hash lookup) ──────────────────────────
    # Map key_string → first available b index
    key_b_to_idx: dict[str, int] = {}
    for idx, k in keys_b.items():
        if k not in key_b_to_idx:
            key_b_to_idx[k] = idx

    match_log: list[dict] = []
    a_to_b:    dict[int, int] = {}    # a_row_idx → b_row_idx
    used_b:    set[int]       = set()
    unmatched_a_idxs: list[int] = []

    for a_idx in range(len(keys_a)):
        k = keys_a.at[a_idx]
        b_idx = key_b_to_idx.get(k)
        if b_idx is not None and b_idx not in used_b:
            a_to_b[a_idx] = b_idx
            used_b.add(b_idx)
            match_log.append(
                {"key_a": k, "key_b": k, "score": 100, "match_type": "exact"}
            )
        else:
            unmatched_a_idxs.append(a_idx)

    unmatched_b_idxs = [i for i in range(len(keys_b)) if i not in used_b]

    # ── Phase 2: Fuzzy matching for unmatched rows ───────────────────────────
    if unmatched_a_idxs and unmatched_b_idxs:
        # Mutable candidate list (items removed as they are matched)
        b_pool_vals: list[str] = [keys_b.at[i] for i in unmatched_b_idxs]
        b_pool_idxs: list[int] = list(unmatched_b_idxs)

        for a_idx in unmatched_a_idxs:
            if not b_pool_vals:
                # No more B candidates
                match_log.append(
                    {"key_a": keys_a.at[a_idx], "key_b": None,
                     "score": 0, "match_type": "unmatched_a"}
                )
                continue

            k_a = keys_a.at[a_idx]

            # Blocking: for large pools, pre-filter by first character
            if len(b_pool_vals) > BLOCK_THRESHOLD:
                prefix = k_a[:1].lower()
                filtered = [
                    (v, i) for v, i in zip(b_pool_vals, b_pool_idxs)
                    if v[:1].lower() == prefix
                ]
                search_vals = [v for v, _ in filtered] or b_pool_vals
                search_idxs = [i for _, i in filtered] or b_pool_idxs
            else:
                search_vals = b_pool_vals
                search_idxs = b_pool_idxs

            # rapidfuzz: weighted ratio handles transpositions + substrings
            result = process.extractOne(
                k_a,
                search_vals,
                scorer=fuzz.WRatio,
                score_cutoff=threshold,
            )

            if result is not None:
                matched_str, score, list_pos = result
                # Map back to the full pool position
                b_idx = search_idxs[list_pos]
                full_pool_pos = b_pool_idxs.index(b_idx)

                a_to_b[a_idx] = b_idx
                used_b.add(b_idx)

                # Remove from pool so this B row can't be reused
                b_pool_vals.pop(full_pool_pos)
                b_pool_idxs.pop(full_pool_pos)

                match_log.append(
                    {"key_a": k_a, "key_b": matched_str,
                     "score": round(float(score), 1), "match_type": "fuzzy"}
                )
            else:
                match_log.append(
                    {"key_a": k_a, "key_b": None,
                     "score": 0, "match_type": "unmatched_a"}
                )
    else:
        # No fuzzy phase needed — log any remaining unmatched A rows
        for a_idx in unmatched_a_idxs:
            match_log.append(
                {"key_a": keys_a.at[a_idx], "key_b": None,
                 "score": 0, "match_type": "unmatched_a"}
            )

    # Log unmatched B rows
    final_used_b = set(a_to_b.values())
    for b_idx in range(len(keys_b)):
        if b_idx not in final_used_b:
            match_log.append(
                {"key_a": None, "key_b": keys_b.at[b_idx],
                 "score": 0, "match_type": "unmatched_b"}
            )

    # ── Reconstruct aligned DataFrames ──────────────────────────────────────
    rows_a, rows_b, statuses = [], [], []

    # 1) Matched pairs (exact + fuzzy), sorted by A index for stability
    for a_idx, b_idx in sorted(a_to_b.items()):
        rows_a.append(df_a.iloc[a_idx].tolist())
        rows_b.append(df_b.iloc[b_idx].tolist())
        statuses.append("matched")

    # 2) Unmatched A rows (B side = empty)
    empty_b_row = [""] * len(df_b.columns)
    for a_idx in sorted(set(range(len(df_a))) - set(a_to_b.keys())):
        rows_a.append(df_a.iloc[a_idx].tolist())
        rows_b.append(empty_b_row)
        statuses.append("only_in_a")

    # 3) Unmatched B rows (A side = empty)
    empty_a_row = [""] * len(df_a.columns)
    for b_idx in sorted(set(range(len(df_b))) - final_used_b):
        rows_a.append(empty_a_row)
        rows_b.append(df_b.iloc[b_idx].tolist())
        statuses.append("only_in_b")

    aligned_a = pd.DataFrame(rows_a, columns=df_a.columns)
    aligned_b = pd.DataFrame(rows_b, columns=df_b.columns)

    aligned_a[STATUS_COL] = statuses
    aligned_b[STATUS_COL] = statuses

    return (
        aligned_a.reset_index(drop=True),
        aligned_b.reset_index(drop=True),
        match_log,
    )


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _make_key(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    """Build a composite key string from one or more columns."""
    if len(cols) == 1:
        return df[cols[0]].fillna("").astype(str)
    return (
        df[cols]
        .fillna("")
        .astype(str)
        .apply(lambda row: _SEP.join(row), axis=1)
    )
