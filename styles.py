"""
CSS + small JS tweaks for the CareLink UI. Kept out of app.py so the
main entry point stays short.
"""
import streamlit as st
import streamlit.components.v1 as components

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* Background */
.stApp {
    background-color: #EDF2F8 !important;
}

header[data-testid="stHeader"] {
    display: none;
}

.block-container {
    padding-top: 1.8rem !important;
    padding-bottom: 2rem !important;
    max-width: 980px !important;
}

/* App Header */
.app-header {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 18px;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 24px;
}

.app-header-logo {
    width: 28px;
    height: 28px;
    object-fit: contain;
    border-radius: 6px;
}

/* Section Titles */
.hero-tag {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 4px;
}
.hero-tag-blue { color: #2F5FE8; }
.hero-tag-green { color: #16213E; }

.hero-title {
    font-size: 32px;
    font-weight: 800;
    color: #0F172A;
    margin-bottom: 6px;
    line-height: 1.1;
    letter-spacing: -0.02em;
}

.hero-subtitle {
    font-size: 14px;
    color: #64748B;
    line-height: 1.5;
    margin-bottom: 24px;
}

/* Card Container Base */
.card-box {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 16px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.02);
}

.card-title {
    font-size: 13.5px;
    color: #334155;
    font-weight: 500;
    margin-bottom: 12px;
}

/* Instruction Tags Grid */
.columns-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    margin-top: 12px;
}
.column-tag {
    background-color: #F8FAFC;
    color: #475569;
    font-size: 12px;
    font-weight: 500;
    padding: 5px 12px;
    border-radius: 6px;
    border: 1px solid #E2E8F0;
}

/* File Upload Display Box */
.uploader-card {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 16px;
}

.uploader-title {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    margin-bottom: 4px;
}

.uploader-sub {
    font-size: 13px;
    color: #475569;
    margin-bottom: 12px;
}

.fake-upload-slot {
    background-color: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 10px 14px;
    display: flex;
    align-items: center;
    gap: 12px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px;
}

/* Monospace Input Styling */
div[data-baseweb="input"] input {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 13px !important;
    color: #1E293B !important;
    background-color: #F8FAFC !important;
    border-radius: 8px !important;
    border: 1px solid #E2E8F0 !important;
    padding-left: 14px !important;
}

/* Status Badges */
.badge-wrapper {
    display: flex;
    align-items: center;
    gap: 8px;
    height: 100%;
    align-items: center;
}

.pill-badge {
    padding: 6px 14px;
    border-radius: 8px;
    font-size: 12px;
    font-weight: 700;
    display: inline-flex;
    align-items: center;
    gap: 6px;
}

.pill-active-blue {
    background-color: #EAF1FE;
    color: #2F5FE8;
    border: 1px solid #C9DBFC;
}

.pill-active-green {
    background-color: #EAECF3;
    color: #16213E;
    border: 1px solid #C7CCDE;
}

.pill-inactive {
    background-color: #F1F5F9;
    color: #94A3B8;
    border: 1px solid #E2E8F0;
}

/* Tabs Styling */
.stTabs [data-baseweb="tab-list"] {
    gap: 24px;
    border-bottom: 1px solid #E2E8F0;
    padding-bottom: 0px;
    margin-bottom: 24px;
}

.stTabs [data-baseweb="tab"] {
    height: 38px;
    background-color: transparent;
    font-weight: 600;
    font-size: 14px;
    color: #64748B;
    border: none;
    border-bottom: 2.5px solid transparent;
    padding: 0 4px 8px 4px;
}

.stTabs [aria-selected="true"] {
    background-color: transparent !important;
    color: #0F172A !important;
    border-bottom: 2.5px solid #16213E !important;
}

/* =========================================================
   FILE UPLOADER — restyled to match the CareLink mock exactly:
   compact single row, no drag/drop copy, "Upload" button,
   filename shown as a blue mono link next to the button.
   ========================================================= */

div[data-testid="stFileUploader"] {
    width: 100%;
}

div[data-testid="stFileUploader"] > section,
div[data-testid="stFileUploaderDropzone"] {
    background-color: #F8FAFC !important;
    border: 1px solid #E2E8F0 !important;
    border-radius: 8px !important;
    padding: 10px 14px !important;
    display: flex !important;
    flex-direction: row !important;
    align-items: center !important;
    gap: 12px !important;
    min-height: unset !important;
}

div[data-testid="stFileUploaderDropzoneInstructions"] {
    display: flex !important;
    align-items: center !important;
    gap: 0 !important;
    order: 3;
}
div[data-testid="stFileUploaderDropzoneInstructions"] svg {
    display: none !important;
}
div[data-testid="stFileUploaderDropzoneInstructions"] span {
    display: none !important;
}
div[data-testid="stFileUploaderDropzoneInstructions"] small {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 11.5px !important;
    color: #94A3B8 !important;
}

div[data-testid="stFileUploaderDropzone"] button {
    background-color: #2F5FE8 !important;
    border: 1px solid #2F5FE8 !important;
    border-radius: 8px !important;
    color: #FFFFFF !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    height: 34px !important;
    padding: 0 14px !important;
    box-shadow: none !important;
    order: 1;
}
div[data-testid="stFileUploaderDropzone"] button:hover {
    background-color: #2650C9 !important;
    border-color: #2650C9 !important;
}
div[data-testid="stFileUploaderDropzone"] button p {
    color: #FFFFFF !important;
}

div[data-testid="stFileUploaderFile"] {
    background: transparent !important;
    border: none !important;
    padding: 0 !important;
    order: 2;
    flex: 1 1 auto;
    min-width: 0;
}
div[data-testid="stFileUploaderFile"] > div:first-child {
    display: none !important;
}
div[data-testid="stFileUploaderFileName"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 13px !important;
    color: #2F5FE8 !important;
    font-weight: 500 !important;
    white-space: normal !important;
    overflow-wrap: anywhere !important;
}
div[data-testid="stFileUploaderFile"] small {
    display: none !important;
}
div[data-testid="stFileUploaderDeleteBtn"] {
    order: 4;
}
div[data-testid="stFileUploaderDeleteBtn"] button svg {
    color: #94A3B8 !important;
}

/* Primary Action Buttons */
div.stButton > button {
    width: 100%;
    height: 42px;
    border-radius: 8px;
    background-color: #16213E !important;
    color: #FFFFFF !important;
    font-size: 14px !important;
    font-weight: 700 !important;
    border: none !important;
    box-shadow: 0 1px 3px rgba(22, 33, 62, 0.35);
    transition: all 0.15s ease;
}

div.stButton > button:hover {
    background-color: #0F1730 !important;
}

div.stButton > button:disabled {
    background-color: #CBD5E1 !important;
    color: #64748B !important;
    box-shadow: none !important;
    cursor: not-allowed;
}

/* Banner */
.success-banner {
    background-color: #EAF1FE;
    border: 1px solid #C9DBFC;
    border-radius: 8px;
    padding: 12px 16px;
    color: #16213E;
    font-size: 13px;
    font-weight: 500;
    margin-top: 16px;
}
</style>
"""


def inject_custom_css():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def inject_uploader_tweaks():
    """
    Streamlit's file_uploader widget has no prop to change its copy
    ("Browse files" -> "Upload", "Limit 200MB..." -> "200MB...").
    This runs a tiny script in the parent document (same-origin iframe
    trick) to rewrite that text so it matches the mock, and re-applies
    itself on every rerun via a MutationObserver since Streamlit
    re-renders the DOM on each interaction.
    """
    components.html(
        """
        <script>
        (function() {
            const doc = window.parent.document;

            function tweak() {
                doc.querySelectorAll('div[data-testid="stFileUploaderDropzone"] button').forEach(btn => {
                    const label = btn.querySelector('div, span, p') || btn;
                    if (label && label.textContent.trim().toLowerCase().includes('browse')) {
                        label.textContent = '↑  Upload';
                    }
                });
                doc.querySelectorAll('div[data-testid="stFileUploaderDropzoneInstructions"] small').forEach(el => {
                    if (el.textContent.startsWith('Limit ')) {
                        el.textContent = el.textContent.replace('Limit ', '');
                    }
                });
            }

            tweak();
            const observer = new MutationObserver(tweak);
            observer.observe(doc.body, { childList: true, subtree: true });
        })();
        </script>
        """,
        height=0,
        width=0,
    )
