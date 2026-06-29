"""
core/loader.py  —  Phase 2
--------------------------
Unified file loader for CSV (.csv), legacy Excel (.xls), and modern Excel (.xlsx).

Key additions over Phase 1
--------------------------
* load_file()           — dispatches to CSV or Excel parser based on extension
* validate_same_type()  — enforces both uploads share the same format family
* get_sheet_names()     — lists sheets for multi-sheet Excel files
* load_csv()            — preserved as a backward-compatible alias
"""

import io
import os
import chardet
import pandas as pd
import openpyxl


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def load_file(file_obj, sheet_name=0) -> tuple[pd.DataFrame, str]:
    """
    Load a CSV, XLS, or XLSX file into a DataFrame.

    Parameters
    ----------
    file_obj   : Streamlit UploadedFile or file-like object
    sheet_name : Sheet index or name for Excel files (default 0 = first sheet)

    Returns
    -------
    df       : pd.DataFrame  — all columns as strings
    info_str : str           — encoding (CSV) or sheet name used (Excel)
    """
    ext = _get_extension(file_obj)
    raw = _read_bytes(file_obj)

    if ext == ".csv":
        return _load_csv_bytes(raw)
    elif ext in (".xlsx", ".xls"):
        return _load_excel_bytes(raw, ext, sheet_name)
    else:
        raise ValueError(
            f"Unsupported file type '{ext}'. "
            "Please upload a .csv, .xls, or .xlsx file."
        )


# Backward-compatible alias used by Phase 1 code paths
def load_csv(file_obj) -> tuple[pd.DataFrame, str]:
    """Alias for load_file() that always treats the input as CSV."""
    raw = _read_bytes(file_obj)
    return _load_csv_bytes(raw)


def validate_same_type(file_a, file_b) -> None:
    """
    Raise ValueError if the two uploaded files belong to different format families.

    CSV (.csv) and Excel (.xls / .xlsx) are treated as two distinct families.
    Mixing XLS and XLSX is *allowed* — both are considered 'Excel'.

    Raises
    ------
    ValueError  with a human-readable explanation.
    """
    family_a = _format_family(_get_extension(file_a))
    family_b = _format_family(_get_extension(file_b))

    if family_a != family_b:
        label_a = family_a.upper()
        label_b = family_b.upper()
        raise ValueError(
            f"❌ File type mismatch — "
            f"**{_safe_name(file_a)}** is {label_a} "
            f"but **{_safe_name(file_b)}** is {label_b}. "
            "Both files must be the same format (both CSV or both Excel)."
        )


def get_sheet_names(file_obj) -> list[str]:
    """
    Return a list of sheet names for an Excel file.
    Returns an empty list for CSV files.
    """
    ext = _get_extension(file_obj)
    if ext not in (".xlsx", ".xls"):
        return []

    raw = _read_bytes(file_obj)

    if ext == ".xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    else:  # .xls
        try:
            import xlrd
            wb = xlrd.open_workbook(file_contents=raw)
            return wb.sheet_names()
        except ImportError:
            return ["Sheet1"]


def preview(df: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    """Return the first *n* rows for display."""
    return df.head(n)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _load_csv_bytes(raw: bytes) -> tuple[pd.DataFrame, str]:
    """Parse raw CSV bytes; auto-detect encoding."""
    detected  = chardet.detect(raw)
    encoding  = detected.get("encoding") or "utf-8"

    try:
        df = pd.read_csv(
            io.BytesIO(raw),
            encoding=encoding,
            dtype=str,
            keep_default_na=False,
        )
    except (UnicodeDecodeError, LookupError):
        df = pd.read_csv(
            io.BytesIO(raw),
            encoding="utf-8",
            errors="replace",
            dtype=str,
            keep_default_na=False,
        )
        encoding = "utf-8 (fallback)"

    df = _clean_df(df)
    return df, encoding


def _load_excel_bytes(raw: bytes, ext: str, sheet_name) -> tuple[pd.DataFrame, str]:
    """Parse raw Excel bytes using openpyxl (.xlsx) or xlrd (.xls)."""
    engine = "openpyxl" if ext == ".xlsx" else "xlrd"

    df = pd.read_excel(
        io.BytesIO(raw),
        sheet_name=sheet_name,
        dtype=str,
        keep_default_na=False,
        engine=engine,
    )
    df = _clean_df(df)

    # info_str = the sheet name that was actually used
    info_str = str(sheet_name) if isinstance(sheet_name, str) else f"Sheet {sheet_name}"
    return df, info_str


def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace from column names and all string cell values."""
    df.columns = [str(c).strip() for c in df.columns]
    df = df.apply(lambda col: col.str.strip() if col.dtype == object else col)
    return df


def _get_extension(file_obj) -> str:
    """Extract the lowercase file extension from a file-like object."""
    name = getattr(file_obj, "name", "") or ""
    _, ext = os.path.splitext(name)
    return ext.lower()


def _format_family(ext: str) -> str:
    """Map an extension to a broad format family name."""
    if ext == ".csv":
        return "csv"
    if ext in (".xlsx", ".xls"):
        return "excel"
    return "unknown"


def _safe_name(file_obj) -> str:
    """Return the filename or a fallback label."""
    return getattr(file_obj, "name", "unknown file")


def _read_bytes(file_obj) -> bytes:
    """Normalise various file-like objects to raw bytes."""
    if isinstance(file_obj, (bytes, bytearray)):
        return bytes(file_obj)
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    return file_obj.read()
