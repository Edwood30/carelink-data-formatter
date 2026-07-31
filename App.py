import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Medicines Formatter",
    page_icon="🧊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================================================
# Custom CSS for Ultra-Clean Modern UI (Matching Image)
# =========================================================
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* Background */
.stApp {
    background-color: #F8FAFC !important;
}

/* Hide default Streamlit padding & header elements */
header[data-testid="stHeader"] {
    display: none;
}
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 2rem !important;
    max-width: 900px !important;
}

/* Navbar */
.nav-bar {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 12px 0 24px 0;
    border-bottom: 1px solid #E2E8F0;
    margin-bottom: 48px;
}
.nav-logo {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 16px;
    font-weight: 700;
    color: #0F172A;
}
.nav-logo svg {
    width: 20px;
    height: 20px;
    fill: #0284C7;
}

/* Hero Section */
.hero-tag {
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 0.08em;
    color: #0284C7;
    text-transform: uppercase;
    margin-bottom: 8px;
}
.hero-title {
    font-size: 42px;
    font-weight: 800;
    color: #0F172A;
    margin-bottom: 12px;
    line-height: 1.1;
    letter-spacing: -0.02em;
}
.hero-subtitle {
    font-size: 16px;
    color: #64748B;
    line-height: 1.5;
    margin-bottom: 36px;
}

/* File Upload Drop Zone */
div[data-testid="stFileUploader"] {
    background: #FFFFFF;
    border: 1.5px dashed #CBD5E1;
    border-radius: 16px;
    padding: 30px 20px;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    transition: border-color 0.2s ease;
}
div[data-testid="stFileUploader"]:hover {
    border-color: #0284C7;
}
div[data-testid="stFileUploader"] section {
    background: transparent !important;
}

/* Download Button Styling */
div.stDownloadButton > button {
    width: 100%;
    height: 48px;
    border-radius: 8px;
    background-color: #0F172A !important;
    color: #FFFFFF !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    border: none !important;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    transition: all 0.2s ease;
}
div.stDownloadButton > button:hover {
    background-color: #1E293B !important;
    transform: translateY(-1px);
}

/* Success Card */
.success-banner {
    background-color: #F0FDF4;
    border: 1px solid #BBF7D0;
    border-radius: 12px;
    padding: 16px 20px;
    color: #166534;
    font-size: 14px;
    font-weight: 500;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 10px;
}
</style>
""",
    unsafe_allow_html=True,
)


def process_rendered_medicines(df):
    """Processes DataFrame and returns openpyxl workbook buffer."""
    for date_col in ["Consultation Date", "Rendered Date"]:
        if date_col in df.columns:
            df[date_col] = (
                pd.to_datetime(df[date_col], errors="coerce")
                .dt.strftime("%m/%d/%Y")
                .fillna("")
            )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rendered Medicines"
    ws.views.sheetView[0].showGridLines = True

    merge_cols = [
        "Patient Name",
        "Last Name",
        "First Name",
        "Middle Name",
        "Patient PIN",
        "Patient Source",
        "Consultation Date",
        "Rendered Date",
        "End Visit By",
        "ICD10 Code",
        "ICD10 Description",
        "Yakap Status",
        "Pharmacy",
        "Contact Number",
        "Address",
        "Notes",
    ]

    headers = list(df.columns)

    # Styles
    header_fill = PatternFill(
        start_color="1F497D", end_color="1F497D", fill_type="solid"
    )
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

    data_font = Font(name="Segoe UI", size=9.5)
    total_font = Font(name="Segoe UI", size=10, bold=True)

    fill_even = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )
    fill_odd = PatternFill(
        start_color="F9FAFB", end_color="F9FAFB", fill_type="solid"
    )

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )
    top_thick_bottom_double = Border(
        top=Side(border_style="thin", color="1F497D"),
        bottom=Side(border_style="double", color="1F497D"),
    )

    align_header = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    align_center_top = Alignment(
        horizontal="center", vertical="top", wrap_text=True
    )
    align_left_top = Alignment(
        horizontal="left", vertical="top", wrap_text=True
    )

    # 1. Header Row
    ws.row_dimensions[1].height = 28
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_header
        cell.border = thin_border

    # 2. Identify Patient Groups
    patient_blocks = []
    current_pin = None
    block_start = 2

    for row_idx, row in df.iterrows():
        pin = (
            str(row["Patient PIN"])
            if ("Patient PIN" in df.columns and pd.notna(row["Patient PIN"]))
            else f"NO_PIN_{row_idx}"
        )
        excel_row = row_idx + 2
        if current_pin is None:
            current_pin = pin
            block_start = excel_row
        elif pin != current_pin:
            patient_blocks.append((block_start, excel_row - 1))
            current_pin = pin
            block_start = excel_row

    if block_start <= len(df) + 1:
        patient_blocks.append((block_start, len(df) + 1))

    # 3. Write Data Rows
    for row_idx, row_data in enumerate(df.values, start=2):
        ws.row_dimensions[row_idx].height = 20
        block_idx = next(
            i for i, (s, e) in enumerate(patient_blocks) if s <= row_idx <= e
        )
        row_fill = fill_even if block_idx % 2 == 0 else fill_odd

        for col_idx, val in enumerate(row_data, start=1):
            col_name = headers[col_idx - 1]
            cell = ws.cell(row=row_idx, column=col_idx)

            cell.value = "" if pd.isna(val) else val
            cell.font = data_font
            cell.fill = row_fill
            cell.border = thin_border

            if col_name in ["Cost", "Price", "Total Cost", "Total Price"]:
                cell.number_format = "₱#,##0.00"
                cell.alignment = align_right
            elif col_name in ["Qty Prescribed", "Qty Dispensed"]:
                cell.number_format = "#,##0"
                cell.alignment = align_right
            elif col_name in [
                "Consultation Date",
                "Rendered Date",
                "Patient PIN",
                "Contact Number",
                "ICD10 Code",
                "Yakap Status",
                "Pharmacy",
            ]:
                cell.alignment = align_center
                if col_name in ["Patient PIN", "Contact Number"]:
                    cell.number_format = "@"
            else:
                cell.alignment = align_left

    # 4. Merge Patient Columns Vertically
    for col_name in merge_cols:
        if col_name in headers:
            col_idx = headers.index(col_name) + 1
            for start_row, end_row in patient_blocks:
                if start_row < end_row:
                    sub_vals = [
                        ws.cell(r, col_idx).value
                        for r in range(start_row, end_row + 1)
                    ]
                    if len(set(sub_vals)) == 1:
                        ws.merge_cells(
                            start_row=start_row,
                            start_column=col_idx,
                            end_row=end_row,
                            end_column=col_idx,
                        )
                        is_centered = col_name in [
                            "Consultation Date",
                            "Rendered Date",
                            "Patient PIN",
                            "Contact Number",
                            "ICD10 Code",
                            "Yakap Status",
                            "Pharmacy",
                        ]
                        target_align = (
                            align_center_top if is_centered else align_left_top
                        )

                        for r in range(start_row, end_row + 1):
                            c = ws.cell(row=r, column=col_idx)
                            c.alignment = target_align
                            c.border = thin_border

    # 5. Total Row
    total_row_idx = len(df) + 2
    ws.row_dimensions[total_row_idx].height = 24

    total_label = ws.cell(row=total_row_idx, column=1, value="Total")
    total_label.font = total_font
    total_label.alignment = Alignment(horizontal="left", vertical="center")

    for col_idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=total_row_idx, column=col_idx)
        cell.font = total_font
        cell.border = top_thick_bottom_double
        col_letter = get_column_letter(col_idx)

        if col_name in ["Qty Prescribed", "Qty Dispensed"]:
            cell.value = f"=SUM({col_letter}2:{col_letter}{total_row_idx-1})"
            cell.number_format = "#,##0"
            cell.alignment = align_right
        elif col_name in ["Total Cost", "Total Price"]:
            cell.value = f"=SUM({col_letter}2:{col_letter}{total_row_idx-1})"
            cell.number_format = "₱#,##0.00"
            cell.alignment = align_right

    # 6. Column Widths
    for col in ws.columns:
        col_idx = col[0].column
        col_name = headers[col_idx - 1]
        max_len = max(
            len(str(cell.value or "")) for cell in col if cell.value is not None
        )
        col_letter = get_column_letter(col_idx)

        if col_name in ["ICD10 Description", "Medicine"]:
            width = min(max(max_len + 3, 25), 45)
        elif col_name in ["Patient Name", "End Visit By"]:
            width = min(max(max_len + 3, 20), 30)
        elif col_name in ["Total Cost", "Total Price", "Cost", "Price"]:
            width = 14
        elif col_name in ["Qty Prescribed", "Qty Dispensed"]:
            width = 15
        else:
            width = min(max(max_len + 4, 12), 25)

        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A2"

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


# =========================================================
# UI Layout
# =========================================================

# Navigation Bar
st.markdown(
    """
<div class="nav-bar">
    <div class="nav-logo">
        <svg viewBox="0 0 24 24"><path d="M4 4h6v6H4V4zm10 0h6v6h-6V4zM4 14h6v6H4v-6zm10 0h6v6h-6v-6z"/></svg>
        Medicines Formatter
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# Header Title Section
st.markdown(
    """
<div class="hero-tag">Rendered Medicines</div>
<div class="hero-title">Report Formatter</div>
<div class="hero-subtitle">
    Convert raw medicine reports into clean, professionally formatted Excel workbooks<br>
    — ready for review in one step.
</div>
""",
    unsafe_allow_html=True,
)

# File Uploader Container
uploaded_file = st.file_uploader(
    "Drop your file here", type=["csv", "xlsx"], help="CSV or Excel · up to 200 MB"
)

# File Processing & Download Action
if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    if uploaded_file.name.endswith(".csv"):
        df_raw = pd.read_csv(io.BytesIO(file_bytes))
    else:
        df_raw = pd.read_excel(io.BytesIO(file_bytes))

    out_name = (
        uploaded_file.name.rsplit(".", 1)[0] + "_Formatted.xlsx"
        if "." in uploaded_file.name
        else "Rendered_Medicines_Formatted.xlsx"
    )

    with st.spinner("Formatting workbook..."):
        excel_buffer = process_rendered_medicines(df_raw)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f"""
    <div class="success-banner">
        <span>✓</span> <strong>{uploaded_file.name}</strong> processed successfully and ready for export.
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.download_button(
        label="Download Formatted Excel Report",
        data=excel_buffer,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )