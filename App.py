"""
CareLink Data Formatting Suite — main entry point.

Run with: streamlit run app.py

Logic lives in separate modules so this file stays short:
- utils.py              shared helpers (column matching, filename cleanup, etc.)
- styles.py             CSS + file-uploader JS tweaks
- feature_formatter.py  Feature 1: Rendered Medicines Report Formatter
- feature_merger.py     Feature 2: Automated Patient & Contact Merger
"""
import base64
import os

import streamlit as st

from feature_formatter import render_formatter_tab
from feature_merger import render_merger_tab
from styles import inject_custom_css, inject_uploader_tweaks

# =========================================================
# LOGO — "FTCC Head.png" is expected to sit alongside this
# script. Drop the file in the same folder as app.py.
# =========================================================
LOGO_FILENAME = "FTCC Head.png"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(APP_DIR, LOGO_FILENAME)


def _find_logo_path():
    """
    Resolves the logo file, tolerating case/spacing differences
    (e.g. "ftcc head.PNG") in case the exact filename doesn't match.
    Returns the real path on disk, or None if nothing matches.
    """
    if os.path.isfile(LOGO_PATH):
        return LOGO_PATH
    try:
        target = LOGO_FILENAME.strip().lower()
        for fname in os.listdir(APP_DIR):
            if fname.strip().lower() == target:
                return os.path.join(APP_DIR, fname)
    except OSError:
        pass
    return None


@st.cache_data(show_spinner=False)
def _load_logo_base64_cached(path, mtime):
    """mtime is part of the cache key so adding/replacing the file while
    the app server is already running doesn't serve a stale cached result."""
    with open(path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("utf-8")
    ext = os.path.splitext(path)[1].lstrip(".").lower() or "png"
    mime = "jpeg" if ext == "jpg" else ext
    return f"data:image/{mime};base64,{encoded}"


def load_logo_base64():
    resolved = _find_logo_path()
    if not resolved:
        return None
    return _load_logo_base64_cached(resolved, os.path.getmtime(resolved))


LOGO_DATA_URI = load_logo_base64()

_resolved_logo_path = _find_logo_path()

# Page Configuration
st.set_page_config(
    page_title="CareLink Data Formatting Suite",
    page_icon=_resolved_logo_path if _resolved_logo_path else "🧩",
    layout="wide",
    initial_sidebar_state="collapsed",
)

inject_custom_css()
inject_uploader_tweaks()

# =========================================================
# APP HEADER & TABS NAVIGATION
# =========================================================
_header_icon_html = (
    f'<img src="{LOGO_DATA_URI}" alt="FTCC" class="app-header-logo" style="width:22px;height:22px;max-width:22px;max-height:22px;object-fit:contain;border-radius:4px;display:block;" />'
    if LOGO_DATA_URI
    else """<svg width="20" height="20" viewBox="0 0 24 24" fill="#2F5FE8">
        <path d="M3 3h8v8H3V3zm10 0h8v8h-8V3zM3 13h8v8H3v-8zm10 0h8v8h-8v-8z"/>
    </svg>"""
)

st.markdown(
    f"""
<div class="app-header-row" style="display:flex;align-items:flex-start;justify-content:space-between;gap:16px;padding-bottom:20px;margin-bottom:24px;border-bottom:1px solid #E2E8F0;">
    <div class="app-header-left" style="display:flex;align-items:center;gap:12px;">
        <div class="app-header-icon-wrap" style="display:flex;align-items:center;justify-content:center;width:34px;height:34px;flex-shrink:0;overflow:hidden;border-radius:9px;background-color:#FFFFFF;border:1px solid #E2E8F0;">{_header_icon_html}</div>
        <div>
            <div class="app-header-title" style="font-size:17px;font-weight:700;color:#0F172A;line-height:1.25;">CareLink Data Formatting Suite</div>
            <div class="app-header-subtitle" style="font-size:12.5px;color:#64748B;line-height:1.4;">Upload CareLink reports, get clean formatted spreadsheets.</div>
        </div>
    </div>
    <span class="pill-badge pill-active-blue app-header-badge" style="flex-shrink:0;margin-top:2px;">CARELINK OPS</span>
</div>
""",
    unsafe_allow_html=True,
)

if not LOGO_DATA_URI:
    with st.expander("⚠️ Logo not loading — click for details"):
        st.markdown(f"Looking for `{LOGO_FILENAME}` in:\n\n`{APP_DIR}`")
        try:
            found = os.listdir(APP_DIR)
            st.markdown("Files actually in that folder:")
            st.code("\n".join(found) if found else "(empty)")
        except OSError as e:
            st.markdown(f"Could not list that folder: {e}")
        st.markdown(
            "If your file is listed above with a different name/case, rename it "
            f"to exactly `{LOGO_FILENAME}`, or place it in the same folder as `app.py` "
            "and restart the app."
        )

tab1, tab2 = st.tabs(["📄 CareLink Express Data", "⊕ CareLink Coordinator Data"])

with tab1:
    render_formatter_tab()

with tab2:
    render_merger_tab()