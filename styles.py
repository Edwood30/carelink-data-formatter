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
    display: flex !important;
    align-items: flex-start !important;
    justify-content: space-between !important;
    gap: 16px;
    padding-bottom: 20px;
    margin-bottom: 24px;
    border-bottom: 1px solid #E2E8F0;
}

.app-header-left {
    display: flex !important;
    align-items: center !important;
    gap: 12px;
}

.app-header-icon-wrap {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 34px;
    height: 34px;
    flex-shrink: 0;
    overflow: hidden;
    border-radius: 9px;
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
}

.app-header-title {
    font-size: 16px;
    font-weight: 700;
    color: #0F172A;
    line-height: 1.25;
}

.app-header-subtitle {
    font-size: 12px;
    color: #64748B;
    line-height: 1.3;
    margin-top: 1px;
}

.app-header-badge {
    flex-shrink: 0;
    margin-top: 2px;
}

.app-header-logo {
    width: 22px !important;
    height: 22px !important;
    max-width: 22px !important;
    max-height: 22px !important;
    object-fit: contain !important;
    border-radius: 4px !important;
    display: block !important;
}

/* Section heading — plain, utilitarian (no oversized marketing type) */
.section-heading {
    font-size: 16px;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 4px;
    letter-spacing: -0.01em;
}

.section-subtext {
    font-size: 13px;
    color: #64748B;
    line-height: 1.5;
    margin-bottom: 20px;
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

.uploader-accent-blue { color: #2F5FE8; }
.uploader-accent-green { color: #16213E; }

/* Outputs-by-source list */
.source-list-header {
    display: flex;
    align-items: center;
    padding: 4px 4px 8px 4px;
    border-bottom: 1px solid #E2E8F0;
    margin-bottom: 4px;
}
.source-list-header span {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    color: #94A3B8;
}

.source-row {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 2px 4px;
}
.source-row-name {
    font-size: 13.5px;
    color: #0F172A;
    font-weight: 500;
    display: inline-flex;
    align-items: center;
    gap: 7px;
}
.source-row-count {
    color: #94A3B8;
    font-weight: 400;
    font-size: 12.5px;
}
.source-icon {
    color: #94A3B8;
    flex-shrink: 0;
    display: inline-flex;
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

/* Tabs Styling — segmented pill toggle.
   NOTE: newer Streamlit versions dropped the old data-baseweb attributes
   on tabs entirely — the real hooks are data-testid="stTabs"/"stTab" plus
   the standard ARIA role/aria-selected attributes. Selectors below are
   written to match both the current and the older markup so this keeps
   working across Streamlit versions. */
[data-testid="stTabs"] [role="tablist"] {
    display: flex;
    gap: 6px;
    background-color: #F1F5F9;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 6px;
    margin-bottom: 24px;
}

[data-testid="stTabs"] [data-testid="stTab"],
[data-testid="stTabs"] [role="tab"] {
    flex: 1;
    height: 42px;
    display: flex !important;
    align-items: center;
    justify-content: center;
    background-color: transparent !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    color: #64748B !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0 8px !important;
    transition: all 0.15s ease;
}

[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"],
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
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

   NOTE: newer Streamlit renders the dropzone as a <section>, not a
   <div>, and renders uploaded files as "stFileChip" elements (not
   "stFileUploaderFile"). Selectors below use bare attribute
   selectors (no leading tag name) so they match regardless of which
   HTML tag Streamlit puts them on, and cover both naming schemes.
   ========================================================= */

[data-testid="stFileUploader"] {
    width: 100%;
}

[data-testid="stFileUploader"] > section,
[data-testid="stFileUploaderDropzone"] {
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
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #94A3B8 !important;
    background-color: #F8FAFC !important;
}

/* Hide Streamlit's own icon/instructions — we render our own below */
[data-testid="stFileUploaderDropzoneInstructions"] {
    display: none !important;
}

/* Browse/Upload button becomes an invisible full-cover click target */
[data-testid="stFileUploaderDropzone"] button {
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

/* Uploaded-file row — current Streamlit calls these "FileChip"s */
[data-testid="stFileChip"] {
    background-color: #F8FAFC !important;
    border: 1px solid #E2E8F0 !important;
    border-radius: 8px !important;
    padding: 8px 12px !important;
    margin-top: 10px !important;
}
[data-testid="stFileChipName"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 13px !important;
    color: #2F5FE8 !important;
    font-weight: 500 !important;
    white-space: normal !important;
    overflow-wrap: anywhere !important;
}
[data-testid="stFileChipDeleteBtn"] svg {
    color: #94A3B8 !important;
}

/* Older Streamlit naming, kept for backward compatibility */
[data-testid="stFileUploaderFile"] {
    background-color: #F8FAFC !important;
    border: 1px solid #E2E8F0 !important;
    border-radius: 8px !important;
    padding: 8px 12px !important;
    margin-top: 10px !important;
}
[data-testid="stFileUploaderFile"] > div:first-child {
    display: none !important;
}
[data-testid="stFileUploaderFileName"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 13px !important;
    color: #2F5FE8 !important;
    font-weight: 500 !important;
    white-space: normal !important;
    overflow-wrap: anywhere !important;
}
[data-testid="stFileUploaderFile"] small {
    display: none !important;
}
[data-testid="stFileUploaderDeleteBtn"] button svg {
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

/* Compact checkboxes for the outputs-by-source list */
[data-testid="stCheckbox"] {
    padding-top: 6px;
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
                // Inject the big centered icon + "Click to upload or drag
                // and drop" visual once per dropzone (guarded so repeated
                // MutationObserver callbacks don't duplicate it). The real
                // Streamlit button stays functionally in place underneath
                // (see CSS: it's made fully transparent and stretched over
                // the whole box), so no button-label rewriting is needed.
                doc.querySelectorAll('[data-testid="stFileUploaderDropzone"]').forEach(zone => {
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