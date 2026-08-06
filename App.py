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
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), LOGO_FILENAME)


@st.cache_data(show_spinner=False)
def load_logo_base64(path):
    """Reads the logo file and returns a base64 data URI, or None if missing."""
    try:
        with open(path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
        return f"data:image/png;base64,{encoded}"
    except FileNotFoundError:
        return None


LOGO_DATA_URI = load_logo_base64(LOGO_PATH)

# Page Configuration
st.set_page_config(
    page_title="CareLink Data Formatting Suite",
    page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "🧩",
    layout="wide",
    initial_sidebar_state="collapsed",
)

inject_custom_css()
inject_uploader_tweaks()

# =========================================================
# APP HEADER & TABS NAVIGATION
# =========================================================
_header_icon_html = (
    f'<img src="{LOGO_DATA_URI}" alt="FTCC" class="app-header-logo" />'
    if LOGO_DATA_URI
    else """<svg width="20" height="20" viewBox="0 0 24 24" fill="#2F5FE8">
        <path d="M3 3h8v8H3V3zm10 0h8v8h-8V3zM3 13h8v8H3v-8zm10 0h8v8h-8v-8z"/>
    </svg>"""
)

st.markdown(
    f"""
<div class="app-header">
    {_header_icon_html}
    CareLink Data Formatting Suite
</div>
""",
    unsafe_allow_html=True,
)

tab1, tab2 = st.tabs(["📄 CareLink Express Data", "⊕ CareLink Coordinator Data"])

with tab1:
    render_formatter_tab()

with tab2:
    render_merger_tab()
