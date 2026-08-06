"""
Shared helper functions used by both CareLink features.
"""
import pandas as pd


def read_data_file(uploaded_file):
    """Helper to read CSV or XLSX file."""
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    else:
        return pd.read_excel(uploaded_file)


def format_contact_number(val):
    """Formats phone number to 11 digits (09xxxxxxxxx)."""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    val_str = str(val).split(".")[0].strip()
    if len(val_str) == 10 and val_str.startswith("9"):
        return "0" + val_str
    return val_str


def find_column(df, candidates, exclude=None):
    """
    Case/whitespace-insensitive column lookup, with a loose contains-match
    fallback. Returns the *actual* column name from df, or None if nothing
    matches any candidate. `exclude` lets a caller rule out columns already
    claimed by another field, so a generic candidate (e.g. "Category")
    can't accidentally re-match a column meant for something else.
    """
    exclude = set(exclude or [])
    normalized = {
        str(c).strip().lower(): c for c in df.columns if c not in exclude
    }
    for cand in candidates:
        key = cand.strip().lower()
        if key in normalized:
            return normalized[key]
    for cand in candidates:
        key = cand.strip().lower()
        for norm_key, orig in normalized.items():
            if key in norm_key or norm_key in key:
                return orig
    return None


def clean_filename(user_input, default_name):
    """Ensures file name ends with .xlsx and is not empty."""
    name = user_input.strip() if user_input else ""
    if not name or name == ".xlsx":
        return default_name
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name
