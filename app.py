"""
app.py  —  Phase 2
-------------------
Streamlit entry-point for the CSV / Excel Validation & Comparison Tool.

How to run
----------
    pip install -r requirements.txt
    streamlit run app.py

Phase 2 additions vs Phase 1
-----------------------------
* Accepts .csv, .xls, and .xlsx files (same-type enforcement)
* Composite Primary Key alignment mode
* Fuzzy alignment (single key & composite key) via rapidfuzz
* Drag-and-drop column reordering with searchable add/remove
* All "File A / File B" labels replaced with actual filenames post-upload
* Excel sheet selector for multi-sheet workbooks
* Fuzzy Match Log tab in Visual Diff + 4th sheet in Excel report
"""

import os
import streamlit as st
import pandas as pd

from core.loader          import load_file, validate_same_type, get_sheet_names, preview
from core.aligner         import align_positional, align_by_key, align_by_composite_key, STATUS_COL
from core.fuzzy_aligner   import align_fuzzy
from core.comparator      import compare, summary_stats
from core.reporter        import to_excel, to_html, to_txt
from components.sortable_mapper import render_sortable_mapper


# ═══════════════════════════════════════════════════════════════════════════════
# Page configuration
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Data Validator",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ═══════════════════════════════════════════════════════════════════════════════
# Global CSS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* ── Hero ── */
    .hero {
        background: linear-gradient(135deg, #0a1628 0%, #0f2544 45%, #0c3b5e 100%);
        border-radius: 18px;
        padding: 44px 52px;
        margin-bottom: 36px;
        box-shadow: 0 12px 40px rgba(0,0,0,0.25);
        position: relative;
        overflow: hidden;
    }
    .hero::after {
        content: '';
        position: absolute;
        top: -40px; right: -40px;
        width: 200px; height: 200px;
        border-radius: 50%;
        background: rgba(255,255,255,0.03);
    }
    .hero h1 {
        color: #ffffff;
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0 0 10px 0;
        letter-spacing: -0.5px;
    }
    .hero p { color: #a8c4e0; font-size: 1rem; margin: 0; }
    .hero .pill {
        display: inline-block;
        background: rgba(255,255,255,0.1);
        color: #d4e9ff;
        border-radius: 20px;
        padding: 3px 12px;
        font-size: 0.75rem;
        font-weight: 500;
        margin: 10px 4px 0 0;
        border: 1px solid rgba(255,255,255,0.15);
    }

    /* ── Step label badges ── */
    .step-label {
        background: linear-gradient(90deg, #1a3a6e, #0f4c75);
        color: #fff;
        font-weight: 600;
        font-size: 0.75rem;
        letter-spacing: 1.8px;
        text-transform: uppercase;
        padding: 5px 16px;
        border-radius: 20px;
        display: inline-block;
        margin-bottom: 6px;
    }

    /* ── Stat cards ── */
    .stat-card {
        background: linear-gradient(135deg, #f0f6ff, #e8f2ff);
        border: 1px solid #c8dcf5;
        border-radius: 14px;
        padding: 22px 16px;
        text-align: center;
        box-shadow: 0 2px 10px rgba(26,58,110,0.06);
    }
    .stat-number { font-size: 2.1rem; font-weight: 700; color: #1a3a6e; }
    .stat-label  {
        font-size: 0.72rem; color: #6b8cad; margin-top: 6px;
        text-transform: uppercase; letter-spacing: 1px;
    }

    /* ── Badges ── */
    .badge-mismatch {
        background: #ffc7ce; color: #9c0006; border-radius: 6px;
        padding: 3px 12px; font-weight: 600; font-size: 0.85rem;
    }
    .badge-match {
        background: #c6efce; color: #276221; border-radius: 6px;
        padding: 3px 12px; font-weight: 600; font-size: 0.85rem;
    }
    .badge-fuzzy {
        background: #ddebf7; color: #1f4e79; border-radius: 6px;
        padding: 3px 12px; font-weight: 600; font-size: 0.85rem;
    }

    /* ── File type pills on upload cards ── */
    .file-type-csv   { color: #276221; background: #c6efce; }
    .file-type-excel { color: #7e6000; background: #ffeb9c; }

    /* ── Section divider ── */
    .section-divider {
        border: none; border-top: 2px solid #e8f0fe; margin: 32px 0;
    }

    /* ── Mode radio clean-up ── */
    div[role="radiogroup"] label {
        border: 1px solid #d0dff5;
        border-radius: 8px;
        padding: 8px 14px;
        margin: 3px 0;
        transition: all 0.15s;
    }
    div[role="radiogroup"] label:hover {
        background: #edf4ff;
        border-color: #1a3a6e;
    }

    /* ── Download buttons ── */
    .stDownloadButton > button {
        width: 100%; border-radius: 10px !important;
        font-weight: 600 !important;
    }

    /* ── Fuzzy score colour in log table ── */
    .score-high { color: #276221; font-weight: 700; }
    .score-mid  { color: #7e6000; font-weight: 600; }
    .score-low  { color: #9c0006; font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════════════════════════════════════
# Hero header
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown(
    """
    <div class="hero">
      <h1>🔍 Data Validator &amp; Comparator</h1>
      <p>Compare CSV and Excel files with intelligent row alignment, fuzzy matching,
         and colour-coded visual reports — now with drag-and-drop column mapping.</p>
      <span class="pill">CSV</span>
      <span class="pill">XLS / XLSX</span>
      <span class="pill">Composite Keys</span>
      <span class="pill">Fuzzy Matching</span>
      <span class="pill">Drag-and-Drop Mapping</span>
    </div>
    """,
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════════════════════════════════════
# Helper: display name for a file (stem without extension, capped at 40 chars)
# ═══════════════════════════════════════════════════════════════════════════════
def _display_name(file_obj) -> str:
    if file_obj is None:
        return "File"
    stem = os.path.splitext(file_obj.name)[0]
    return stem[:40] if len(stem) > 40 else stem


def _ext(file_obj) -> str:
    if file_obj is None:
        return ""
    return os.path.splitext(file_obj.name)[1].upper().lstrip(".")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Upload Files
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<span class="step-label">Step 1</span>', unsafe_allow_html=True)
st.subheader("Upload Files")
st.caption("Supported formats: .csv  ·  .xls  ·  .xlsx  — both files must be the **same format**.")

use_sample = st.checkbox("💡 Use Sample Data (file_a.csv vs file_b.csv)")

if use_sample:
    class LocalFileWrapper:
        def __init__(self, filepath):
            self.filepath = filepath
            self.name = os.path.basename(filepath)
        def read(self):
            with open(self.filepath, "rb") as f:
                return f.read()
        def seek(self, offset, whence=0):
            pass
    file_a = LocalFileWrapper("sample_data/file_a.csv")
    file_b = LocalFileWrapper("sample_data/file_b.csv")
else:
    up1, up2 = st.columns(2)

    with up1:
        st.markdown("**📄 Reference File** *(source of truth)*")
        file_a = st.file_uploader(
            "Upload Reference File",
            type=["csv", "xls", "xlsx"],
            key="upload_a",
            label_visibility="collapsed",
        )

    with up2:
        st.markdown("**📄 Comparison File** *(to validate against reference)*")
        file_b = st.file_uploader(
            "Upload Comparison File",
            type=["csv", "xls", "xlsx"],
            key="upload_b",
            label_visibility="collapsed",
        )

if not file_a or not file_b:
    st.info("👆 Upload **both** files above or check 'Use Sample Data' to continue.")
    st.stop()

# ── Extract display names immediately after upload ───────────────────────────
name_a = _display_name(file_a)
name_b = _display_name(file_b)
ext_a  = _ext(file_a)
ext_b  = _ext(file_b)

# ── Enforce same-type constraint ─────────────────────────────────────────────
try:
    validate_same_type(file_a, file_b)
except ValueError as e:
    st.error(str(e))
    st.stop()

# ── Sheet selector for multi-sheet Excel files ───────────────────────────────
sheet_a = 0
sheet_b = 0

sheets_a = get_sheet_names(file_a)
sheets_b = get_sheet_names(file_b)

if sheets_a or sheets_b:
    sh1, sh2 = st.columns(2)
    if sheets_a:
        with sh1:
            sheet_a = st.selectbox(
                f"Sheet in **{name_a}**",
                options=sheets_a,
                key="sheet_a",
            )
    if sheets_b:
        with sh2:
            sheet_b = st.selectbox(
                f"Sheet in **{name_b}**",
                options=sheets_b,
                key="sheet_b",
            )

# ── Load files ────────────────────────────────────────────────────────────────
try:
    df_a_raw, info_a = load_file(file_a, sheet_name=sheet_a)
    df_b_raw, info_b = load_file(file_b, sheet_name=sheet_b)
except Exception as e:
    st.error(f"❌ Failed to load files: {e}")
    st.stop()

# ── File info strip ───────────────────────────────────────────────────────────
ic1, ic2 = st.columns(2)
with ic1:
    st.success(
        f"**{name_a}** ({ext_a}) — {len(df_a_raw):,} rows · "
        f"{len(df_a_raw.columns)} columns · `{info_a}`"
    )
with ic2:
    st.success(
        f"**{name_b}** ({ext_b}) — {len(df_b_raw):,} rows · "
        f"{len(df_b_raw.columns)} columns · `{info_b}`"
    )

# ── File preview ─────────────────────────────────────────────────────────────
with st.expander("👁️ Preview first 10 rows", expanded=False):
    pv1, pv2 = st.columns(2)
    with pv1:
        st.markdown(f"**{name_a}**")
        st.dataframe(preview(df_a_raw), use_container_width=True)
    with pv2:
        st.markdown(f"**{name_b}**")
        st.dataframe(preview(df_b_raw), use_container_width=True)

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Column Mapping (drag-and-drop + searchable)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<span class="step-label">Step 2</span>', unsafe_allow_html=True)
st.subheader("Column Mapping")
st.markdown(
    f"Map columns from **{name_a}** (File A) to columns from **{name_b}** (File B). "
    "Expand **Column Selection** to add/remove columns. "
    "Use the **dropdown** on each row to change which File B column is paired. "
    "Use **↑ ↓** buttons to reorder pairs."
)

cols_a = df_a_raw.columns.tolist()
cols_b = df_b_raw.columns.tolist()

selected_a, selected_b = render_sortable_mapper(cols_a, cols_b, name_a, name_b)

# Validate selection
if not selected_a or not selected_b:
    st.warning("⚠️ Select at least one column from each file to continue.")
    st.stop()

if len(selected_a) != len(selected_b):
    st.error(
        f"❌ Column count mismatch — **{len(selected_a)}** selected from "
        f"**{name_a}** and **{len(selected_b)}** from **{name_b}**. "
        "The counts must be equal."
    )
    st.stop()

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Comparison Options
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<span class="step-label">Step 3</span>', unsafe_allow_html=True)
st.subheader("Comparison Options")

# ── Alignment mode ────────────────────────────────────────────────────────────
ALIGN_MODES = [
    "Positional  (row 1 → row 1)",
    "Single Key — Exact Match",
    "Composite Key — Exact Match",
    "Single Key — Fuzzy Match",
    "Composite Key — Fuzzy Match",
]

align_mode = st.radio(
    "Row alignment mode",
    options=ALIGN_MODES,
    index=0,
    horizontal=False,
    help=(
        "**Positional**: compare rows by position.  "
        "**Single Key**: join rows by one shared column (exact).  "
        "**Composite Key**: join rows by 2+ columns combined (exact).  "
        "**Fuzzy Single/Composite**: like key-based but tolerates typos/variations."
    ),
)

# ── Key column selectors ──────────────────────────────────────────────────────
key_cols:     list[str] = []
fuzzy_threshold: int    = 80

is_single_key    = "Single Key" in align_mode
is_composite_key = "Composite Key" in align_mode
is_fuzzy         = "Fuzzy" in align_mode

if is_single_key:
    kc1, kc2 = st.columns([1, 2])
    with kc1:
        key_col = st.selectbox(
            f"Key column (from **{name_a}** / shared name after mapping)",
            options=selected_a,
            help="Values in this column uniquely identify each row. "
                 "The same logical column is used in both files.",
        )
    key_cols = [key_col]

elif is_composite_key:
    kc1, kc2 = st.columns([2, 1])
    with kc1:
        key_cols = st.multiselect(
            "Composite key columns (select 2 or more)",
            options=selected_a,
            default=selected_a[:2] if len(selected_a) >= 2 else selected_a,
            help="The combination of these columns uniquely identifies each row.",
        )
    if len(key_cols) < 2:
        st.warning("⚠️ Please select **at least 2 columns** for a composite key.")
        st.stop()

# ── Fuzzy threshold ───────────────────────────────────────────────────────────
if is_fuzzy:
    tc1, tc2 = st.columns([1, 2])
    with tc1:
        fuzzy_threshold = st.slider(
            "Fuzzy match threshold (0 – 100)",
            min_value=50,
            max_value=100,
            value=80,
            step=1,
            help=(
                "Minimum similarity score (rapidfuzz WRatio) required to accept "
                "a fuzzy match.  **80** is a safe default.  "
                "Lower = more lenient (more matches, risk of bad pairings).  "
                "Higher = stricter (fewer matches, higher quality)."
            ),
        )
    with tc2:
        st.info(
            f"🔵 A score ≥ **{fuzzy_threshold}** is required to pair two rows.  "
            "Rows below the threshold are flagged as unmatched."
        )

# ── Other options ─────────────────────────────────────────────────────────────
oc1, oc2 = st.columns(2)
with oc1:
    tolerance = st.number_input(
        "Numeric tolerance",
        min_value=0.0, max_value=1000.0, value=0.0, step=0.001, format="%.4f",
        help="If > 0, numeric cells within this absolute difference are treated "
             "as matching.  Set to 0 for strict equality.",
    )
with oc2:
    case_sensitive = st.toggle(
        "Case-sensitive text comparison",
        value=True,
        help="When OFF, 'Hello' and 'hello' are treated as equal.",
    )

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Run Comparison
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<span class="step-label">Step 4</span>', unsafe_allow_html=True)
st.subheader("Run Comparison")

run_btn = st.button("▶️ Run Comparison", type="primary", use_container_width=True)

if not run_btn:
    st.info(f"Click **▶️ Run Comparison** to compare **{name_a}** against **{name_b}**.")
    st.stop()


# ── Normalise (case) ──────────────────────────────────────────────────────────
def _normalise(df: pd.DataFrame) -> pd.DataFrame:
    if case_sensitive:
        return df
    return df.apply(lambda col: col.str.lower() if col.dtype == object else col)


# ── Extract mapped columns ────────────────────────────────────────────────────
df_a_mapped = _normalise(df_a_raw[selected_a].copy())
df_b_mapped = _normalise(df_b_raw[selected_b].copy())

# Rename B columns to match A's names for unified comparison
rename_map  = dict(zip(selected_b, selected_a))
df_b_mapped = df_b_mapped.rename(columns=rename_map)


# ── Row alignment (dispatch to correct aligner) ───────────────────────────────
match_log: list[dict] | None = None   # populated only in fuzzy modes

with st.spinner(f"Aligning rows ({align_mode.split('—')[0].strip()})…"):

    if align_mode.startswith("Positional"):
        aligned_a, aligned_b = align_positional(df_a_mapped, df_b_mapped)

    elif align_mode.startswith("Single Key — Exact"):
        aligned_a, aligned_b = align_by_key(
            df_a_mapped, df_b_mapped, key_col=key_cols[0]
        )

    elif align_mode.startswith("Composite Key — Exact"):
        aligned_a, aligned_b = align_by_composite_key(
            df_a_mapped, df_b_mapped, key_cols=key_cols
        )

    elif align_mode.startswith("Single Key — Fuzzy"):
        aligned_a, aligned_b, match_log = align_fuzzy(
            df_a_mapped, df_b_mapped,
            key_cols=key_cols,
            threshold=fuzzy_threshold,
        )

    elif align_mode.startswith("Composite Key — Fuzzy"):
        aligned_a, aligned_b, match_log = align_fuzzy(
            df_a_mapped, df_b_mapped,
            key_cols=key_cols,
            threshold=fuzzy_threshold,
        )

    else:
        aligned_a, aligned_b = align_positional(df_a_mapped, df_b_mapped)


# ── Cell-level comparison ──────────────────────────────────────────────────
with st.spinner("Comparing cells…"):
    diff_mask, mismatch_list = compare(aligned_a, aligned_b, tolerance=tolerance)
    stats = summary_stats(diff_mask, aligned_a)

# ── Enrich mismatch_list with primary-key values ─────────────────────────────
# This adds a 'primary_key' field to each mismatch dict, used by the
# Mismatch Detail Excel sheet and the in-app table.
for m in mismatch_list:
    row_idx = m["row"] - 1  # convert to 0-based
    if key_cols and 0 <= row_idx < len(aligned_a):
        pk_parts = [
            f"{kc}={aligned_a.at[row_idx, kc]}"
            for kc in key_cols
            if kc in aligned_a.columns
        ]
        m["primary_key"] = " | ".join(pk_parts) if pk_parts else "N/A"
    else:
        m["primary_key"] = "N/A"

# Human-readable label for the primary-key column header
if not key_cols:
    pk_label = "Primary Key"
elif len(key_cols) == 1:
    pk_label = f"Primary Key ({key_cols[0]})"
else:
    pk_label = "Composite Key (" + " + ".join(key_cols) + ")"


# ═══════════════════════════════════════════════════════════════════════════════
# Results — Stats banner
# ═══════════════════════════════════════════════════════════════════════════════
sc1, sc2, sc3, sc4 = st.columns(4)
for widget, label, value in [
    (sc1, "Total Rows",       stats["total_rows"]),
    (sc2, "Total Cells",      stats["total_cells"]),
    (sc3, "Mismatched Cells", stats["mismatched_cells"]),
    (sc4, "Match Rate",       f"{stats['match_rate_pct']}%"),
]:
    widget.markdown(
        f'<div class="stat-card">'
        f'<div class="stat-number">{value}</div>'
        f'<div class="stat-label">{label}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

st.markdown("")   # spacing

# Fuzzy summary pill
if match_log:
    n_exact  = sum(1 for m in match_log if m["match_type"] == "exact")
    n_fuzzy  = sum(1 for m in match_log if m["match_type"] == "fuzzy")
    n_unmatch = sum(1 for m in match_log if "unmatched" in m["match_type"])
    st.markdown(
        f'<span class="badge-match">✅ {n_exact} exact</span> &nbsp;'
        f'<span class="badge-fuzzy">🔵 {n_fuzzy} fuzzy</span> &nbsp;'
        f'<span class="badge-mismatch">⚠️ {n_unmatch} unmatched</span>',
        unsafe_allow_html=True,
    )
    st.markdown("")

if stats["cols_with_diffs"]:
    st.markdown(
        f"**Columns with mismatches in {name_a} vs {name_b}:** "
        + ", ".join(
            f'<span class="badge-mismatch">{c}</span>'
            for c in stats["cols_with_diffs"]
        ),
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        f'<span class="badge-match">✅ Perfect match — '
        f'{name_a} and {name_b} are identical for the selected columns!</span>',
        unsafe_allow_html=True,
    )

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# Visual Diff Preview
# ═══════════════════════════════════════════════════════════════════════════════
st.subheader("📊 Visual Diff Preview")
st.caption("🔴 Red = mismatch  ·  🟢 Green = match  ·  🟡 Amber = row-status column")

compare_cols = [c for c in diff_mask.columns if c != STATUS_COL]

def _highlight_diff(row):
    """Pandas Styler: colour each cell from diff_mask."""
    styles = []
    row_idx = row.name
    for col in row.index:
        if col == STATUS_COL:
            styles.append("background-color:#fff8dc;color:#555;")
        elif col in compare_cols and row_idx in diff_mask.index and col in diff_mask.columns:
            if diff_mask.at[row_idx, col]:
                styles.append("background-color:#ffc7ce;color:#9c0006;font-weight:600;")
            else:
                styles.append("background-color:#c6efce;color:#276221;")
        else:
            styles.append("")
    return styles


# Build tabs: File A | File B | [Fuzzy Log]
tab_labels = [f"📄 {name_a}", f"📄 {name_b}"]
if match_log:
    tab_labels.append("🔵 Fuzzy Match Log")

tabs = st.tabs(tab_labels)

with tabs[0]:
    styled_a = aligned_a.style.apply(_highlight_diff, axis=1)
    st.dataframe(styled_a, use_container_width=True, height=420)

with tabs[1]:
    styled_b = aligned_b.style.apply(_highlight_diff, axis=1)
    st.dataframe(styled_b, use_container_width=True, height=420)

if match_log and len(tabs) > 2:
    with tabs[2]:
        st.caption(
            "Each row shows how a key from one file was matched to the other. "
            "🟢 Exact  ·  🔵 Fuzzy  ·  🔴 Unmatched"
        )
        log_df = pd.DataFrame(match_log)
        log_df.columns = [
            f"Key ({name_a})", f"Key ({name_b})", "Score", "Match Type"
        ]

        def _colour_log_row(row):
            mt = row["Match Type"]
            if mt == "exact":
                bg = "background-color:#c6efce;color:#276221;"
            elif mt == "fuzzy":
                bg = "background-color:#ddebf7;color:#1f4e79;"
            else:
                bg = "background-color:#ffc7ce;color:#9c0006;"
            return [bg] * len(row)

        st.dataframe(
            log_df.style.apply(_colour_log_row, axis=1),
            use_container_width=True,
            height=400,
        )

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# Download Reports
# ═══════════════════════════════════════════════════════════════════════════════
st.subheader("⬇️ Download Reports")
st.caption(f"All reports are labelled with the actual filenames: **{name_a}** and **{name_b}**.")

with st.spinner("Generating reports…"):
    xlsx_bytes = to_excel(
        aligned_a, aligned_b, diff_mask, stats,
        name_a=name_a, name_b=name_b,
        match_log=match_log,
        mismatch_list=mismatch_list,
        pk_label=pk_label,
    )
    html_str   = to_html(aligned_a, aligned_b, diff_mask, stats,
                         name_a=name_a, name_b=name_b)
    txt_str    = to_txt(mismatch_list, stats,
                        name_a=name_a, name_b=name_b, match_log=match_log)

# Dynamic download filename from actual file stems
dl_stem = f"{name_a}_vs_{name_b}"[:80]

dl1, dl2, dl3 = st.columns(3)

with dl1:
    st.download_button(
        label="📥 Download Excel Report (.xlsx)",
        data=xlsx_bytes,
        file_name=f"{dl_stem}_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with dl2:
    st.download_button(
        label="🌐 Download HTML Report (.html)",
        data=html_str.encode("utf-8"),
        file_name=f"{dl_stem}_report.html",
        mime="text/html",
        use_container_width=True,
    )

with dl3:
    st.download_button(
        label="📄 Download Text Summary (.txt)",
        data=txt_str.encode("utf-8"),
        file_name=f"{dl_stem}_summary.txt",
        mime="text/plain",
        use_container_width=True,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# Mismatch Detail Table
# ═══════════════════════════════════════════════════════════════════════════════
if mismatch_list:
    st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)
    st.subheader("🔎 Mismatch Detail")
    st.caption(
        f"Every cell that differs between **{name_a}** and **{name_b}**.  "
        f"The **{pk_label}** column shows the record identifier for each row."
    )
    mismatch_df = pd.DataFrame([
        {
            "Row #":       m.get("row", ""),
            "Column":      m.get("column", ""),
            pk_label:      m.get("primary_key", "N/A"),
            f"Value in {name_a}": m.get("value_a", ""),
            f"Value in {name_b}": m.get("value_b", ""),
        }
        for m in mismatch_list
    ])
    # Highlight the PK column teal and the value columns red
    def _style_mismatch_df(df: pd.DataFrame) -> pd.DataFrame:
        styles = pd.DataFrame("", index=df.index, columns=df.columns)
        pk_col = pk_label
        if pk_col in df.columns:
            styles[pk_col] = "background-color:#d9ead3;color:#274e13;font-weight:600;"
        for c in [f"Value in {name_a}", f"Value in {name_b}"]:
            if c in df.columns:
                styles[c] = "background-color:#ffcccc;color:#9c0006;font-weight:600;"
        return styles
    st.dataframe(
        mismatch_df.style.apply(_style_mismatch_df, axis=None),
        use_container_width=True,
        hide_index=True,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# Footer
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown(
    '<p style="text-align:center;color:#bbb;font-size:0.72rem;margin-top:52px;">'
    "Data Validator &amp; Comparator — Phase 2 · Streamlit + Python · rapidfuzz"
    "</p>",
    unsafe_allow_html=True,
)
