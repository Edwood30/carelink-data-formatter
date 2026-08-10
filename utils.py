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
    # Word-boundary fuzzy fallback: only match when every word in the
    # candidate we're searching for also appears in the actual column name
    # (e.g. "PIN" -> "Patient PIN"). Matching in the reverse direction is
    # deliberately NOT done — e.g. searching for "Medicine Category" must
    # not match a column simply named "Medicine", since "medicine" being a
    # substring/word-subset of "medicine category" would otherwise steal
    # that column away from the real "Medicine" search.
    for cand in candidates:
        cand_words = set(cand.strip().lower().split())
        for norm_key, orig in normalized.items():
            norm_words = set(norm_key.split())
            if cand_words and cand_words <= norm_words:
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


def clean_str(val):
    """NaN-safe strip; also collapses internal double-spaces some exports have."""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    while "  " in s:
        s = s.replace("  ", " ")
    return s


def clean_pin(val):
    """Strips a PIN and removes the trailing '.0' Excel/pandas sometimes adds
    when a numeric-looking ID column gets read in as a float."""
    s = clean_str(val)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def clean_date(val):
    """Parses any recognizable date value and returns it as MM/DD/YYYY, or ''
    if it can't be parsed (blank/garbage input)."""
    parsed = pd.to_datetime(val, errors="coerce")
    return parsed.strftime("%m/%d/%Y") if pd.notna(parsed) else ""


def clean_number(val):
    """Coerces a value to a plain number for qty/cost/price columns; returns
    '' (blank cell) if it isn't numeric at all."""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    try:
        n = float(str(val).replace(",", "").strip())
        return int(n) if n == int(n) else n
    except (ValueError, TypeError):
        return ""


def split_name_fallback(name):
    """Best-effort split of a single combined name field into Last/First/Middle."""
    name = clean_str(name)
    if not name:
        return "", "", ""
    if "," in name:
        last, rest = name.split(",", 1)
        rest_parts = rest.strip().split()
        first = rest_parts[0] if rest_parts else ""
        middle = " ".join(rest_parts[1:]) if len(rest_parts) > 1 else ""
        return last.strip(), first, middle
    parts = name.split()
    if len(parts) == 1:
        return "", parts[0], ""
    first = parts[0]
    last = parts[-1]
    middle = " ".join(parts[1:-1])
    return last, first, middle