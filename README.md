# CSV Validator & Comparator

A production-grade Python tool to **compare two CSV files cell-by-cell**, highlight discrepancies, and generate Excel, HTML, and text reports — all through a clean Streamlit browser UI.

---

## Features

| Feature | Details |
|---|---|
| **Positional matching** | Row 1 of File A vs Row 1 of File B |
| **Key-based matching** | Match rows by a shared ID column (e.g. `order_id`) |
| **Column mapping** | Select any subset of columns from each file to compare |
| **Numeric tolerance** | Treat `1.001` and `1.002` as equal when tolerance ≥ 0.001 |
| **Case-insensitive mode** | Optional toggle for text comparisons |
| **Auto encoding detection** | UTF-8, UTF-16, Latin-1 — handled automatically |
| **Missing row detection** | Shorter files are padded; extra rows are flagged |
| **Excel report** | 3-sheet `.xlsx` with green/red cell colouring + Summary sheet |
| **HTML report** | Self-contained `.html` with inline CSS colour table |
| **Text summary** | Detailed `.txt` listing every mismatch by row/column/value |

---

## Project Structure

```
CsvValidator/
├── app.py                  # Streamlit UI entry-point
├── requirements.txt        # Python dependencies
├── core/
│   ├── __init__.py
│   ├── loader.py           # CSV loading + encoding detection
│   ├── aligner.py          # Positional & key-based row alignment
│   ├── comparator.py       # Cell-level diff engine
│   └── reporter.py         # Excel / HTML / TXT report generators
└── sample_data/
    ├── file_a.csv          # Demo reference file (10 rows, 10 cols)
    └── file_b.csv          # Demo comparison file (6 intentional mismatches)
```

---

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the App

```bash
streamlit run app.py
```

Your browser will open at `http://localhost:8501`.

### 3. Use the 4-Step UI

| Step | What to do |
|---|---|
| **Step 1 — Upload** | Upload your two CSV files |
| **Step 2 — Column Mapping** | Select which columns from each file to compare |
| **Step 3 — Options** | Choose alignment mode, tolerance, case sensitivity |
| **Step 4 — Run** | Click **▶️ Run Comparison** and download reports |

---

## Sample Data

Use the files in `sample_data/` to try the tool immediately.  
`file_b.csv` has **6 intentional mismatches** vs `file_a.csv`:

| Row | Column | File A | File B |
|---|---|---|---|
| 2 | status | `Shipped` | `Dispatched` |
| 3 | discount_pct | `0.10` | `0.15` |
| 5 | quantity | `3` | `4` |
| 8 | region | `West` | `North` |
| 9 | unit_price | `19.99` | `20.00` |
| 10 | total_price | `299.99` | `349.99` |

---

## Output Files

| File | Description |
|---|---|
| `csv_comparison_report.xlsx` | 3-sheet Excel: File A (coloured), File B (coloured), Summary stats |
| `csv_comparison_report.html` | Self-contained HTML with colour-coded tables |
| `csv_comparison_summary.txt` | Plain-text mismatch list with row/column/value details |

---

## Dependencies

```
pandas>=2.0.0
openpyxl>=3.1.0
streamlit>=1.35.0
chardet>=5.2.0
```
