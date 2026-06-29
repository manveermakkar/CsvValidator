"""
core/reporter.py
----------------
Generates three output formats from the comparison results:

1. Excel (.xlsx)  — openpyxl with red/green cell fills + header row
2. HTML  (.html)  — self-contained, inline-CSS colored table
3. Text  (.txt)   — human-readable mismatch summary

All three functions accept an ``io.BytesIO`` / ``io.StringIO`` as ``dest``
so they work seamlessly both with Streamlit's download_button (in-memory)
and with ordinary file paths (write to disk).
"""

import io
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from core.aligner import STATUS_COL
from core.comparator import summary_stats


# ── Colour constants ─────────────────────────────────────────────────────────
_GREEN_FILL  = PatternFill("solid", fgColor="C6EFCE")   # light green
_RED_FILL    = PatternFill("solid", fgColor="FFC7CE")    # light red
_YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")    # amber (status/header)
_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")    # dark navy header
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def to_excel(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    diff_mask: pd.DataFrame,
    stats: dict,
) -> bytes:
    """
    Build an Excel workbook in memory and return raw bytes.

    Workbook contains three sheets:
    - ``File A``       — all rows from File A, cells coloured green/red
    - ``File B``       — all rows from File B, cells coloured green/red
    - ``Summary``      — key statistics and per-column mismatch counts
    """
    wb = openpyxl.Workbook()

    # ── Sheet: File A ──────────────────────────────────────────────────────
    ws_a = wb.active
    ws_a.title = "File A"
    _write_coloured_sheet(ws_a, df_a, diff_mask, label="A")

    # ── Sheet: File B ──────────────────────────────────────────────────────
    ws_b = wb.create_sheet("File B")
    _write_coloured_sheet(ws_b, df_b, diff_mask, label="B")

    # ── Sheet: Summary ─────────────────────────────────────────────────────
    ws_s = wb.create_sheet("Summary")
    _write_summary_sheet(ws_s, stats, diff_mask)

    # Serialise to bytes
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def to_html(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    diff_mask: pd.DataFrame,
    stats: dict,
) -> str:
    """
    Build a self-contained HTML string with inline CSS.

    Returns the full HTML document as a string (UTF-8).
    """
    compare_cols = [c for c in diff_mask.columns if c != STATUS_COL]

    def _cell_style(row_idx: int, col: str, is_a: bool) -> str:
        """Return an inline style string for a single data cell."""
        if col == STATUS_COL:
            return "background:#fff8dc; color:#555; font-style:italic;"
        if col not in compare_cols:
            return ""
        mismatch = diff_mask.at[row_idx, col] if col in diff_mask.columns else False
        if mismatch:
            label = "A" if is_a else "B"
            return f"background:#ffc7ce; color:#9c0006; font-weight:600;" \
                   f" title='Mismatch in {col} (row {row_idx+1}, side {label})'"
        return "background:#c6efce; color:#276221;"

    def _build_table(df: pd.DataFrame, title: str, is_a: bool) -> str:
        cols = df.columns.tolist()
        rows_html = []
        for i, row in df.iterrows():
            cells = "".join(
                f'<td style="padding:6px 10px; border:1px solid #ccc; '
                f'{_cell_style(i, col, is_a)}">'
                f"{_esc(str(row[col]))}</td>"
                for col in cols
            )
            row_num = f'<td style="padding:6px 10px; border:1px solid #ccc; ' \
                      f'background:#f0f0f0; color:#666; font-weight:600;">' \
                      f"{i + 1}</td>"
            rows_html.append(f"<tr>{row_num}{cells}</tr>")

        header_cells = "".join(
            f'<th style="padding:8px 12px; background:#1f4e79; color:#fff; '
            f'border:1px solid #0d3457; text-align:left;">{_esc(c)}</th>'
            for c in cols
        )
        header = (
            f'<th style="padding:8px 12px; background:#0d3457; color:#fff; '
            f'border:1px solid #0d3457;">#</th>'
            + header_cells
        )

        return (
            f'<h2 style="color:#1f4e79; margin-top:40px; font-family:sans-serif;">'
            f"{title}</h2>"
            f'<div style="overflow-x:auto;">'
            f'<table style="border-collapse:collapse; font-family:monospace; '
            f'font-size:13px; min-width:600px;">'
            f"<thead><tr>{header}</tr></thead>"
            f"<tbody>{''.join(rows_html)}</tbody>"
            f"</table></div>"
        )

    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    table_a = _build_table(df_a, "📄 File A", is_a=True)
    table_b = _build_table(df_b, "📄 File B", is_a=False)

    # Stats banner
    stats_html = (
        f'<div style="display:flex; gap:24px; flex-wrap:wrap; margin:24px 0;">'
        + "".join(
            f'<div style="background:#f4f8ff; border:1px solid #c0d4f0; '
            f'border-radius:8px; padding:16px 24px; text-align:center;">'
            f'<div style="font-size:28px; font-weight:700; color:#1f4e79;">'
            f"{v}</div>"
            f'<div style="color:#555; font-size:12px; margin-top:4px;">{k}</div>'
            f"</div>"
            for k, v in {
                "Total Rows":       stats["total_rows"],
                "Total Cells":      stats["total_cells"],
                "Mismatched Cells": stats["mismatched_cells"],
                "Match Rate":       f"{stats['match_rate_pct']}%",
            }.items()
        )
        + "</div>"
    )

    cols_with_diffs = stats.get("cols_with_diffs", [])
    diff_cols_html = (
        f'<p style="font-family:sans-serif; color:#9c0006;">'
        f"<strong>Columns with mismatches:</strong> "
        f"{', '.join(cols_with_diffs) if cols_with_diffs else 'None'}</p>"
    )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>CSV Comparison Report</title>
  <style>
    body {{ margin: 40px; background: #f9fafb; }}
    h1   {{ color: #1f4e79; font-family: sans-serif; }}
    p    {{ font-family: sans-serif; color: #444; }}
  </style>
</head>
<body>
  <h1>📊 CSV Comparison Report</h1>
  <p style="color:#888;">Generated: {ts}</p>
  {stats_html}
  {diff_cols_html}
  <hr style="border:none; border-top:2px solid #ddd; margin:32px 0;"/>
  {table_a}
  {table_b}
  <p style="color:#aaa; font-size:11px; margin-top:40px;">
    Generated by CSV Validator Tool
  </p>
</body>
</html>"""
    return html


def to_txt(mismatch_list: list[dict], stats: dict) -> str:
    """
    Generate a detailed plain-text mismatch summary.

    Returns the full report as a string (UTF-8).
    """
    lines: list[str] = []
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    lines += [
        "=" * 72,
        "  CSV COMPARISON REPORT",
        f"  Generated : {ts}",
        "=" * 72,
        "",
        "SUMMARY",
        "-------",
        f"  Total rows compared : {stats['total_rows']}",
        f"  Total cells compared: {stats['total_cells']}",
        f"  Mismatched cells    : {stats['mismatched_cells']}",
        f"  Match rate          : {stats['match_rate_pct']}%",
        f"  Columns with diffs  : "
        + (", ".join(stats["cols_with_diffs"]) if stats["cols_with_diffs"] else "None"),
        "",
        "=" * 72,
        "MISMATCH DETAILS",
        "=" * 72,
        "",
    ]

    if not mismatch_list:
        lines.append("  ✅  No mismatches found — files are identical for the selected columns.")
    else:
        for m in mismatch_list:
            lines += [
                f"  Row    : {m['row']}",
                f"  Column : {m['column']}",
                f"  File A : {m['value_a']!r}",
                f"  File B : {m['value_b']!r}",
                "-" * 48,
            ]

    lines += [
        "",
        "=" * 72,
        "END OF REPORT",
        "=" * 72,
    ]

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _write_coloured_sheet(
    ws,
    df: pd.DataFrame,
    diff_mask: pd.DataFrame,
    label: str,
) -> None:
    """Write a single coloured data sheet into an openpyxl worksheet."""
    compare_cols = [c for c in diff_mask.columns if c != STATUS_COL]

    # ── Header row ─────────────────────────────────────────────────────────
    # Column 1 is a row-number column
    ws.cell(row=1, column=1, value="#").fill = _HEADER_FILL
    ws.cell(row=1, column=1).font = _HEADER_FONT
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")

    for col_idx, col_name in enumerate(df.columns, start=2):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="left")
        cell.border = _THIN_BORDER

    # ── Data rows ──────────────────────────────────────────────────────────
    for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        data_row_num = row_idx - 1   # 1-based display row number

        # Row-number column
        rn_cell = ws.cell(row=row_idx, column=1, value=data_row_num)
        rn_cell.alignment = Alignment(horizontal="center")
        rn_cell.fill = PatternFill("solid", fgColor="E2EFDA")
        rn_cell.border = _THIN_BORDER

        for col_idx, col_name in enumerate(df.columns, start=2):
            value = row[col_name]
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = _THIN_BORDER
            cell.alignment = Alignment(wrap_text=True)

            # Colour logic
            if col_name == STATUS_COL:
                cell.fill = _YELLOW_FILL
            elif col_name in compare_cols:
                # row_idx-2 because the df has 0-based index but ws starts at row 2
                df_row_idx = row_idx - 2
                is_mismatch = (
                    diff_mask.at[df_row_idx, col_name]
                    if df_row_idx in diff_mask.index and col_name in diff_mask.columns
                    else False
                )
                cell.fill = _RED_FILL if is_mismatch else _GREEN_FILL

    # ── Auto-size columns (capped at 50 chars) ─────────────────────────────
    for col_cells in ws.columns:
        max_len = max(
            (len(str(c.value)) for c in col_cells if c.value is not None),
            default=8,
        )
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max_len + 4, 54)

    # Freeze the header row
    ws.freeze_panes = "A2"


def _write_summary_sheet(ws, stats: dict, diff_mask: pd.DataFrame) -> None:
    """Write the Summary statistics sheet."""
    ws.title = "Summary"

    rows = [
        ("Metric", "Value"),
        ("Total Rows Compared",  stats["total_rows"]),
        ("Total Cells Compared", stats["total_cells"]),
        ("Mismatched Cells",     stats["mismatched_cells"]),
        ("Match Rate (%)",       stats["match_rate_pct"]),
        ("", ""),
        ("Column", "Mismatch Count"),
    ]

    compare_cols = [c for c in diff_mask.columns if c != STATUS_COL]
    for col in compare_cols:
        rows.append((col, int(diff_mask[col].sum())))

    for r_idx, (label, value) in enumerate(rows, start=1):
        cell_l = ws.cell(row=r_idx, column=1, value=label)
        cell_v = ws.cell(row=r_idx, column=2, value=value)
        if r_idx in (1, 7):   # header rows
            cell_l.fill = _HEADER_FILL
            cell_v.fill = _HEADER_FILL
            cell_l.font = _HEADER_FONT
            cell_v.font = _HEADER_FONT
        cell_l.border = _THIN_BORDER
        cell_v.border = _THIN_BORDER

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 20


def _esc(text: str) -> str:
    """HTML-escape a string for safe embedding in HTML."""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
    )
