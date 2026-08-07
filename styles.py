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
.app-header-row {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 16px;
    padding-bottom: 20px;
    margin-bottom: 24px;
    border-bottom: 1px solid #E2E8F0;
}

.app-header-left {
    display: flex;
    align-items: center;
    gap: 12px;
}

.app-header-icon-wrap {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 34px;
    height: 34px;
    flex-shrink: 0;
    border-radius: 9px;
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
}

.app-header-title {
    font-size: 17px;
    font-weight: 700;
    color: #0F172A;
    line-height: 1.25;
}

.app-header-subtitle {
    font-size: 12.5px;
    color: #64748B;
    line-height: 1.4;
}

.app-header-badge {
    flex-shrink: 0;
    margin-top: 2px;
}

.app-header-logo {
    width: 22px;
    height: 22px;
    object-fit: contain;
    border-radius: 4px;
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

/* Upload label above a dropzone ("Upload a [X] file") */
.upload-label {
    font-size: 13.5px;
    font-weight: 600;
    color: #0F172A;
    margin-bottom: 10px;
}
.upload-label-accent {
    color: #2F5FE8;
    font-weight: 700;
}

/* Tabs Styling — segmented pill toggle */
.stTabs [data-baseweb="tab-list"] {
    display: flex;
    gap: 6px;
    background-color: #F1F5F9;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 6px;
    margin-bottom: 24px;
}

.stTabs [data-baseweb="tab"] {
    flex: 1;
    height: 42px;
    display: flex;
    align-items: center;
    justify-content: center;
    background-color: transparent;
    font-weight: 600;
    font-size: 14px;
    color: #64748B;
    border: none;
    border-radius: 8px;
    padding: 0 8px;
    transition: all 0.15s ease;
}

.stTabs [aria-selected="true"] {
    background-color: #16213E !important;
    color: #FFFFFF !important;
    border-radius: 8px !important;
    box-shadow: 0 1px 3px rgba(22, 33, 62, 0.35);
}

/* =========================================================
   FILE UPLOADER — big dashed dropzone with a centered icon and
   "Click to upload or drag and drop" copy (Mission-Report-Parser
   style). The default Streamlit button is kept but made invisible
   and stretched over the whole box so the entire area stays
   clickable; the visible icon/text is injected by JS below.
   ========================================================= */

div[data-testid="stFileUploader"] {
    width: 100%;
}

div[data-testid="stFileUploader"] > section,
div[data-testid="stFileUploaderDropzone"] {
    position: relative !important;
    background-color: #FAFBFC !important;
    border: 2px dashed #CBD5E1 !important;
    border-radius: 12px !important;
    padding: 30px 20px !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    gap: 10px !important;
    min-height: 150px !important;
    transition: all 0.15s ease;
}
div[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #94A3B8 !important;
    background-color: #F8FAFC !important;
}

/* Hide Streamlit's own icon/instructions — we render our own below */
div[data-testid="stFileUploaderDropzoneInstructions"] {
    display: none !important;
}

/* Browse button becomes an invisible full-cover click target */
div[data-testid="stFileUploaderDropzone"] button {
    position: absolute !important;
    inset: 0 !important;
    width: 100% !important;
    height: 100% !important;
    min-height: unset !important;
    opacity: 0 !important;
    cursor: pointer !important;
    border-radius: 12px !important;
    padding: 0 !important;
    margin: 0 !important;
}

/* Our injected icon + "Click to upload or drag and drop" text */
.clk-upload-visual {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 10px;
    pointer-events: none;
}
.clk-upload-icon {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 40px;
    height: 40px;
    border-radius: 50%;
    background-color: #EEF2F7;
    color: #16213E;
}
.clk-upload-text {
    font-size: 13.5px;
    color: #64748B;
    text-align: center;
}
.clk-upload-text strong {
    color: #16213E;
    font-weight: 700;
}

div[data-testid="stFileUploaderFile"] {
    background-color: #F8FAFC !important;
    border: 1px solid #E2E8F0 !important;
    border-radius: 8px !important;
    padding: 8px 12px !important;
    margin-top: 10px !important;
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
                        label.textContent = 'Browse';
                    }
                });

                // Inject the big centered icon + "Click to upload or drag
                // and drop" visual once per dropzone (guarded so repeated
                // MutationObserver callbacks don't duplicate it).
                doc.querySelectorAll('div[data-testid="stFileUploaderDropzone"]').forEach(zone => {
                    if (zone.querySelector('.clk-upload-visual')) return;
                    const visual = doc.createElement('div');
                    visual.className = 'clk-upload-visual';
                    visual.innerHTML = `
                        <div class="clk-upload-icon">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M12 3v12"></path>
                                <path d="M7 8l5-5 5 5"></path>
                                <path d="M5 21h14"></path>
                            </svg>
                        </div>
                        <div class="clk-upload-text"><strong>Click to upload</strong> or drag and drop</div>
                    `;
                    zone.appendChild(visual);
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