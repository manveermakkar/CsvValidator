"""
app.py
------
Streamlit entry-point for the CSV Validation & Comparison Tool.

How to run
----------
    pip install -r requirements.txt
    streamlit run app.py

The UI guides the user through four steps:
  1. Upload two CSV files
  2. Map columns (select which columns from each file to compare)
  3. Configure options (alignment mode, numeric tolerance)
  4. Run the comparison and download the three output reports
"""

import io
import streamlit as st
import pandas as pd

from core.loader     import load_csv, preview
from core.aligner    import align_positional, align_by_key, STATUS_COL
from core.comparator import compare, summary_stats
from core.reporter   import to_excel, to_html, to_txt


# ── Page configuration ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="CSV Validator",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* Dark gradient hero header */
    .hero {
        background: linear-gradient(135deg, #0f2544 0%, #1a3a6e 50%, #0f4c75 100%);
        border-radius: 16px;
        padding: 40px 48px;
        margin-bottom: 32px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.18);
    }
    .hero h1 {
        color: #ffffff;
        font-size: 2.4rem;
        font-weight: 700;
        margin: 0 0 8px 0;
        letter-spacing: -0.5px;
    }
    .hero p {
        color: #a8c4e0;
        font-size: 1rem;
        margin: 0;
    }

    /* Step labels */
    .step-label {
        background: linear-gradient(90deg, #1a3a6e, #0f4c75);
        color: #fff;
        font-weight: 600;
        font-size: 0.8rem;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        padding: 4px 14px;
        border-radius: 20px;
        display: inline-block;
        margin-bottom: 8px;
    }

    /* Stat cards */
    .stat-card {
        background: #f0f6ff;
        border: 1px solid #c8dcf5;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
    }
    .stat-number {
        font-size: 2rem;
        font-weight: 700;
        color: #1a3a6e;
    }
    .stat-label {
        font-size: 0.78rem;
        color: #6b8cad;
        margin-top: 4px;
        text-transform: uppercase;
        letter-spacing: 0.8px;
    }

    /* Mismatch badge */
    .badge-mismatch { background:#ffc7ce; color:#9c0006; border-radius:6px;
                      padding:2px 10px; font-weight:600; font-size:0.85rem; }
    .badge-match    { background:#c6efce; color:#276221; border-radius:6px;
                      padding:2px 10px; font-weight:600; font-size:0.85rem; }

    /* Section divider */
    .section-divider {
        border: none;
        border-top: 2px solid #e8f0fe;
        margin: 28px 0;
    }

    /* Download buttons row */
    .stDownloadButton > button {
        width: 100%;
        border-radius: 10px !important;
        font-weight: 600 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ── Hero header ──────────────────────────────────────────────────────────────
st.markdown(
    """
    <div class="hero">
      <h1>🔍 CSV Validator & Comparator</h1>
      <p>Upload two CSV files, map their columns, and get a colour-coded
         Excel report, an HTML report, and a plain-text mismatch summary — in seconds.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Upload
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<span class="step-label">Step 1</span>', unsafe_allow_html=True)
st.subheader("Upload CSV Files")

col_up1, col_up2 = st.columns(2)

with col_up1:
    st.markdown("**📄 File A** *(reference / source)*")
    file_a = st.file_uploader("Upload File A", type=["csv"], key="file_a",
                               label_visibility="collapsed")

with col_up2:
    st.markdown("**📄 File B** *(comparison target)*")
    file_b = st.file_uploader("Upload File B", type=["csv"], key="file_b",
                               label_visibility="collapsed")

# Stop here until both files are uploaded
if not file_a or not file_b:
    st.info("👆 Please upload **both** CSV files to continue.")
    st.stop()

# Load the CSVs
df_a_raw, enc_a = load_csv(file_a)
df_b_raw, enc_b = load_csv(file_b)

st.caption(
    f"File A: **{file_a.name}** — {len(df_a_raw)} rows, "
    f"{len(df_a_raw.columns)} columns — encoding: `{enc_a}`   |   "
    f"File B: **{file_b.name}** — {len(df_b_raw)} rows, "
    f"{len(df_b_raw.columns)} columns — encoding: `{enc_b}`"
)

# Preview
with st.expander("👁️ Preview first 10 rows", expanded=False):
    pv1, pv2 = st.columns(2)
    with pv1:
        st.markdown(f"**File A** — `{file_a.name}`")
        st.dataframe(preview(df_a_raw), use_container_width=True)
    with pv2:
        st.markdown(f"**File B** — `{file_b.name}`")
        st.dataframe(preview(df_b_raw), use_container_width=True)

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Column Mapping
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<span class="step-label">Step 2</span>', unsafe_allow_html=True)
st.subheader("Column Mapping")
st.markdown(
    "Select which columns from **File A** to compare and map them to the "
    "corresponding columns in **File B**.  "
    "The order of your selection matters — column *i* from A is compared with "
    "column *i* from B."
)

cols_a = df_a_raw.columns.tolist()
cols_b = df_b_raw.columns.tolist()

map_col1, map_col2 = st.columns(2)

with map_col1:
    selected_a = st.multiselect(
        "Columns from **File A** to compare",
        options=cols_a,
        default=cols_a,            # default: all columns selected
        key="selected_a",
        help="Pick the columns from File A you want to include in the comparison.",
    )

with map_col2:
    selected_b = st.multiselect(
        "Columns from **File B** (must match count)",
        options=cols_b,
        default=[c for c in cols_b if c in [
            # pre-select matching names
            col for col in selected_a if col in cols_b
        ]] or cols_b[:len(selected_a)],
        key="selected_b",
        help="Pick the same number of columns from File B, in the same comparison order.",
    )

# Validate column-count parity
if len(selected_a) == 0 or len(selected_b) == 0:
    st.warning("⚠️ Please select at least one column from each file.")
    st.stop()

if len(selected_a) != len(selected_b):
    st.error(
        f"❌ Column count mismatch: you selected **{len(selected_a)}** from File A "
        f"and **{len(selected_b)}** from File B.  They must be equal."
    )
    st.stop()

# Show the mapping as a table
mapping_df = pd.DataFrame({
    "File A Column": selected_a,
    "↔": ["↔"] * len(selected_a),
    "File B Column": selected_b,
})
with st.expander("📋 View column mapping", expanded=True):
    st.dataframe(mapping_df, use_container_width=True, hide_index=True)

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Options
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<span class="step-label">Step 3</span>', unsafe_allow_html=True)
st.subheader("Comparison Options")

opt_col1, opt_col2, opt_col3 = st.columns(3)

with opt_col1:
    alignment_mode = st.radio(
        "Row alignment mode",
        options=["Positional (row 1 → row 1)", "Key-based (join on a column)"],
        index=0,
        help="Positional: compare rows by their position. "
             "Key-based: match rows using a shared column value (e.g. order_id).",
    )

key_col_a = None
key_col_b = None

if alignment_mode.startswith("Key"):
    k1, k2 = st.columns(2)
    with k1:
        key_col_a = st.selectbox(
            "Key column in File A",
            options=selected_a,
            help="This column's values are used to match rows from File A to File B.",
        )
    with k2:
        key_col_b = st.selectbox(
            "Key column in File B",
            options=selected_b,
            help="This column's values are used to match rows from File B to File A.",
        )

with opt_col2:
    tolerance = st.number_input(
        "Numeric tolerance",
        min_value=0.0,
        max_value=1000.0,
        value=0.0,
        step=0.001,
        format="%.4f",
        help="If > 0, numeric cells within this absolute difference are treated as matching. "
             "Set to 0 for strict equality.",
    )

with opt_col3:
    case_sensitive = st.toggle(
        "Case-sensitive text comparison",
        value=True,
        help="When OFF, 'Hello' and 'hello' are treated as equal.",
    )

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Run & Download
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<span class="step-label">Step 4</span>', unsafe_allow_html=True)
st.subheader("Run Comparison")

run_btn = st.button("▶️ Run Comparison", type="primary", use_container_width=True)

if not run_btn:
    st.info("Click **▶️ Run Comparison** to start.")
    st.stop()


# ── Apply case-insensitive option ────────────────────────────────────────────
def _normalise(df: pd.DataFrame) -> pd.DataFrame:
    """Lower-case all string cells when case-insensitive mode is active."""
    if case_sensitive:
        return df
    return df.apply(lambda col: col.str.lower() if col.dtype == object else col)


# ── Extract mapped columns ───────────────────────────────────────────────────
df_a_mapped = _normalise(df_a_raw[selected_a].copy())
df_b_mapped = _normalise(df_b_raw[selected_b].copy())

# Rename File B columns to match File A's names for side-by-side comparison
rename_map = dict(zip(selected_b, selected_a))
df_b_mapped = df_b_mapped.rename(columns=rename_map)


# ── Align rows ───────────────────────────────────────────────────────────────
with st.spinner("Aligning rows…"):
    if alignment_mode.startswith("Key") and key_col_a and key_col_b:
        # For key-based: rename the key column in B to match A's name first
        aligned_a, aligned_b = align_by_key(df_a_mapped, df_b_mapped, key_col=key_col_a)
    else:
        aligned_a, aligned_b = align_positional(df_a_mapped, df_b_mapped)


# ── Compare ──────────────────────────────────────────────────────────────────
with st.spinner("Comparing cells…"):
    diff_mask, mismatch_list = compare(aligned_a, aligned_b, tolerance=tolerance)
    stats = summary_stats(diff_mask, aligned_a)


# ── Results banner ───────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
for col_widget, label, value in [
    (c1, "Total Rows",       stats["total_rows"]),
    (c2, "Total Cells",      stats["total_cells"]),
    (c3, "Mismatched Cells", stats["mismatched_cells"]),
    (c4, "Match Rate",       f"{stats['match_rate_pct']}%"),
]:
    col_widget.markdown(
        f'<div class="stat-card">'
        f'<div class="stat-number">{value}</div>'
        f'<div class="stat-label">{label}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

st.markdown("")   # spacing

if stats["cols_with_diffs"]:
    st.markdown(
        "**Columns with mismatches:** "
        + ", ".join(
            f'<span class="badge-mismatch">{c}</span>'
            for c in stats["cols_with_diffs"]
        ),
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<span class="badge-match">✅ All selected columns match perfectly!</span>',
        unsafe_allow_html=True,
    )

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ── Interactive diff table ───────────────────────────────────────────────────
st.subheader("📊 Visual Diff Preview")
st.caption("Showing the aligned data with mismatches highlighted in 🔴 red.")

compare_cols = [c for c in diff_mask.columns if c != STATUS_COL]

def _highlight_diff(row):
    """Pandas Styler function: colour each cell based on diff_mask."""
    styles = []
    row_idx = row.name
    for col in row.index:
        if col == STATUS_COL:
            styles.append("background-color: #fff8dc; color: #555;")
        elif col in compare_cols and row_idx in diff_mask.index and col in diff_mask.columns:
            if diff_mask.at[row_idx, col]:
                styles.append("background-color: #ffc7ce; color: #9c0006; font-weight: 600;")
            else:
                styles.append("background-color: #c6efce; color: #276221;")
        else:
            styles.append("")
    return styles


tab_a, tab_b = st.tabs([f"📄 File A — {file_a.name}", f"📄 File B — {file_b.name}"])

with tab_a:
    styled_a = aligned_a.style.apply(_highlight_diff, axis=1)
    st.dataframe(styled_a, use_container_width=True, height=400)

with tab_b:
    styled_b = aligned_b.style.apply(_highlight_diff, axis=1)
    st.dataframe(styled_b, use_container_width=True, height=400)

st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)


# ── Generate & download reports ──────────────────────────────────────────────
st.subheader("⬇️ Download Reports")

with st.spinner("Generating reports…"):
    xlsx_bytes = to_excel(aligned_a, aligned_b, diff_mask, stats)
    html_str   = to_html(aligned_a, aligned_b, diff_mask, stats)
    txt_str    = to_txt(mismatch_list, stats)

dl1, dl2, dl3 = st.columns(3)

with dl1:
    st.download_button(
        label="📥 Download Excel Report (.xlsx)",
        data=xlsx_bytes,
        file_name="csv_comparison_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with dl2:
    st.download_button(
        label="🌐 Download HTML Report (.html)",
        data=html_str.encode("utf-8"),
        file_name="csv_comparison_report.html",
        mime="text/html",
        use_container_width=True,
    )

with dl3:
    st.download_button(
        label="📄 Download Text Summary (.txt)",
        data=txt_str.encode("utf-8"),
        file_name="csv_comparison_summary.txt",
        mime="text/plain",
        use_container_width=True,
    )


# ── Mismatch detail table ────────────────────────────────────────────────────
if mismatch_list:
    st.markdown('<hr class="section-divider"/>', unsafe_allow_html=True)
    st.subheader("🔎 Mismatch Detail")
    mismatch_df = pd.DataFrame(mismatch_list)
    mismatch_df.columns = ["Row #", "Column", "Value in File A", "Value in File B"]
    st.dataframe(mismatch_df, use_container_width=True, hide_index=True)


# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown(
    '<p style="text-align:center; color:#aaa; font-size:0.75rem; margin-top:48px;">'
    "CSV Validator &amp; Comparator — built with Streamlit &amp; Python"
    "</p>",
    unsafe_allow_html=True,
)
