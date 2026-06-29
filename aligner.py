"""
core/aligner.py
---------------
Responsible for aligning two DataFrames so that rows that should be compared
end up in the same position.

Two modes are supported:

1. **Positional** (default)
   The Nth row of File A is compared with the Nth row of File B.
   If the files have different numbers of rows, the shorter one is padded
   with empty strings so that "extra" rows in the longer file are surfaced
   as mismatches rather than silently dropped.

2. **Key-based**
   A user-chosen column (e.g. ``order_id``) is used as a primary key.
   Rows are joined on that key using an *outer* merge, so rows present in
   only one file are included and clearly flagged with a ``__row_status__``
   column (``matched`` / ``only_in_a`` / ``only_in_b``).
"""

import pandas as pd
import numpy as np


# Column injected during key-based alignment to flag row provenance.
STATUS_COL = "__row_status__"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def align_positional(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Align two DataFrames by row position.

    The shorter DataFrame is padded with empty-string rows so both outputs
    have the same length.  The caller's column selection has already been
    applied before this function is called, so both DataFrames share the
    same column set.

    Returns
    -------
    aligned_a, aligned_b : pd.DataFrame, pd.DataFrame
        Both DataFrames with identical shape and a clean integer index.
    """
    max_len = max(len(df_a), len(df_b))

    # Pad whichever DF is shorter
    aligned_a = _pad_to_length(df_a, max_len)
    aligned_b = _pad_to_length(df_b, max_len)

    # Reset index so rows are labelled 0, 1, 2, …
    aligned_a = aligned_a.reset_index(drop=True)
    aligned_b = aligned_b.reset_index(drop=True)

    return aligned_a, aligned_b


def align_by_key(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    key_col: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Align two DataFrames by matching on a shared key column.

    Rows present in only one file get empty strings for all columns from the
    other file, and a ``__row_status__`` column is injected:
      - ``'matched'``   — key found in both files
      - ``'only_in_a'`` — key only in File A
      - ``'only_in_b'`` — key only in File B

    Returns
    -------
    aligned_a, aligned_b : pd.DataFrame, pd.DataFrame
        Both DataFrames ordered by the key column, with status information
        carried in ``__row_status__``.
    """
    # Outer merge on the key column.
    # suffixes keep column names distinct inside the merged frame.
    merged = pd.merge(
        df_a.add_suffix("__A"),
        df_b.add_suffix("__B"),
        left_on=f"{key_col}__A",
        right_on=f"{key_col}__B",
        how="outer",
    )

    # Determine the column lists (excluding the key suffix duplicates)
    cols_a = [c for c in df_a.columns]
    cols_b = [c for c in df_b.columns]

    # Rebuild two aligned DataFrames from the merged result
    def _extract(cols, suffix):
        frame = pd.DataFrame(index=merged.index)
        for col in cols:
            src = f"{col}{suffix}"
            frame[col] = merged[src].fillna("") if src in merged.columns else ""
        return frame

    out_a = _extract(cols_a, "__A")
    out_b = _extract(cols_b, "__B")

    # Inject row-status column into both frames
    status = _compute_status(merged, key_col)
    out_a[STATUS_COL] = status
    out_b[STATUS_COL] = status

    # Sort by the key column for a deterministic display order
    sort_key = out_a[key_col] if key_col in out_a.columns else out_a.index
    order = sort_key.argsort()
    out_a = out_a.iloc[order].reset_index(drop=True)
    out_b = out_b.iloc[order].reset_index(drop=True)

    return out_a, out_b


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _pad_to_length(df: pd.DataFrame, target_len: int) -> pd.DataFrame:
    """Append empty-string rows until `df` has exactly `target_len` rows."""
    shortage = target_len - len(df)
    if shortage <= 0:
        return df.copy()

    empty_rows = pd.DataFrame(
        [[""] * len(df.columns)] * shortage,
        columns=df.columns,
    )
    return pd.concat([df, empty_rows], ignore_index=True)


def _compute_status(merged: pd.DataFrame, key_col: str) -> pd.Series:
    """
    Derive per-row status from the outer-merged DataFrame.

    A row is:
    - ``'only_in_a'`` if the B-side key is NaN  (no match in File B)
    - ``'only_in_b'`` if the A-side key is NaN  (no match in File A)
    - ``'matched'``   otherwise
    """
    a_key = merged.get(f"{key_col}__A")
    b_key = merged.get(f"{key_col}__B")

    conditions = [
        b_key.isna(),  # present only in A
        a_key.isna(),  # present only in B
    ]
    choices = ["only_in_a", "only_in_b"]

    return pd.Series(
        np.select(conditions, choices, default="matched"),
        index=merged.index,
    )
