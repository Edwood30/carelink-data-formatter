"""
FEATURE 1: Rendered Medicines Report Formatter.

Takes a single raw "Rendered Medicines" export and turns it into a
professionally formatted Excel workbook (merged patient blocks,
totals row, currency/qty formatting, etc).
"""
import io

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from utils import clean_filename, find_column

MERGE_COLS = [
    "Patient Name",
    "Last Name",
    "First Name",
    "Middle Name",
    "Suffix Name",
    "Patient PIN",
    "Patient Source",
    "Consultation Date",
    "Rendered Date",
    "End Visit By",
    "Dispensed By",
    "ICD10 Code",
    "ICD10 Description",
    "Yakap Status",
    "Pharmacy",
    "Contact Number",
    "Cellphone Number",
    "Address",
    "Full Address",
    "Street Name",
    "Barangay",
    "Municipality",
    "Province",
    "Notes",
]

CENTERED_COLS = [
    "Consultation Date",
    "Rendered Date",
    "Patient PIN",
    "Contact Number",
    "Cellphone Number",
    "ICD10 Code",
    "Yakap Status",
    "Pharmacy",
]


def _sort_by_last_name(df):
    """
    Sorts rows alphabetically by Last Name (then First Name), keeping every
    patient's own rows (e.g. multiple medicines) contiguous — required for
    the patient-block merging further down. Falls back to a combined
    Patient Name column if there's no separate Last/First Name.
    """
    df = df.copy()
    last_col = find_column(df, ["Last Name"])
    first_col = find_column(df, ["First Name"])
    pin_col = find_column(df, ["Patient PIN", "PIN"])
    name_col = find_column(df, ["Patient Name", "Full Name"])

    if last_col:
        df["_sort_last"] = df[last_col].astype(str).str.strip().str.lower()
    elif name_col:
        df["_sort_last"] = df[name_col].astype(str).str.strip().str.lower()
    else:
        df["_sort_last"] = ""

    df["_sort_first"] = (
        df[first_col].astype(str).str.strip().str.lower() if first_col else ""
    )
    df["_sort_pin"] = df[pin_col].astype(str) if pin_col else ""

    df = df.sort_values(
        by=["_sort_last", "_sort_first", "_sort_pin"], kind="stable"
    ).reset_index(drop=True)
    return df.drop(columns=["_sort_last", "_sort_first", "_sort_pin"])


def process_rendered_medicines(df):
    df = _sort_by_last_name(df)

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

    headers = list(df.columns)

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
    align_right = Alignment(horizontal="right", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_center_top = Alignment(
        horizontal="center", vertical="top", wrap_text=True
    )
    align_left_top = Alignment(
        horizontal="left", vertical="top", wrap_text=True
    )

    ws.row_dimensions[1].height = 28
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_header
        cell.border = thin_border

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
            elif col_name in CENTERED_COLS:
                cell.alignment = align_center
                if col_name in ["Patient PIN", "Contact Number", "Cellphone Number"]:
                    cell.number_format = "@"
            else:
                cell.alignment = align_left

    for col_name in MERGE_COLS:
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
                        is_centered = col_name in CENTERED_COLS
                        target_align = (
                            align_center_top if is_centered else align_left_top
                        )

                        for r in range(start_row, end_row + 1):
                            c = ws.cell(row=r, column=col_idx)
                            c.alignment = target_align
                            c.border = thin_border

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


def render_formatter_tab():
    """Draws the full 'CareLink Express Data' tab UI and wires up processing."""
    st.markdown(
        """
    <div class="hero-tag hero-tag-blue">RENDERED MEDICINES</div>
    <div class="hero-title">Report Formatter</div>
    <div class="hero-subtitle">Convert raw medicine reports into clean, professionally formatted Excel workbooks.</div>
    """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
    <div class="card-box">
        <div style="display: flex; align-items: center; gap: 8px; font-weight: 700; font-size: 14px; color: #0F172A;">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="#2F5FE8"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14zM7 7h10v2H7V7zm0 4h10v2H7v-2zm0 4h7v2H7v-2z"/></svg>
            Required Excel (.xlsx) Column Structure
        </div>
        <div style="font-size: 13px; color: #64748B; margin-top: 4px;">
            Ensure your uploaded spreadsheet contains the following standard headers:
        </div>
        <div class="columns-grid">
            <span class="column-tag">Patient Name</span>
            <span class="column-tag">Patient PIN</span>
            <span class="column-tag">Cellphone Number</span>
            <span class="column-tag">Full Address</span>
            <span class="column-tag">Patient Source</span>
            <span class="column-tag">Consultation Date</span>
            <span class="column-tag">Rendered Date</span>
            <span class="column-tag">End Visit By</span>
            <span class="column-tag">Dispensed By</span>
            <span class="column-tag">ICD10 Code</span>
            <span class="column-tag">ICD10 Description</span>
            <span class="column-tag">Medicine</span>
            <span class="column-tag">Medicine Category</span>
            <span class="column-tag">Yakap Status</span>
            <span class="column-tag">Pharmacy</span>
            <span class="column-tag">Qty Prescribed</span>
            <span class="column-tag">Qty Dispensed</span>
            <span class="column-tag">Cost</span>
            <span class="column-tag">Price</span>
            <span class="column-tag">Total Cost</span>
            <span class="column-tag">Total Price</span>
            <span class="column-tag">First Name</span>
            <span class="column-tag">Middle Name</span>
            <span class="column-tag">Last Name</span>
            <span class="column-tag">Suffix Name</span>
            <span class="column-tag">Street Name</span>
            <span class="column-tag">Barangay</span>
            <span class="column-tag">Municipality</span>
            <span class="column-tag">Province</span>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    with st.container():
        st.markdown(
            '<div class="card-title" style="margin-left: 2px;">Drop your file here</div>',
            unsafe_allow_html=True,
        )
        uploaded_file = st.file_uploader(
            "Drop your file here",
            type=["xlsx"],
            key="single_fmt",
            label_visibility="collapsed",
        )

    st.markdown("<br>", unsafe_allow_html=True)

    default_out = "CareLinkExpress-Data-Source.xlsx"

    with st.container():
        st.markdown(
            '<div class="card-title" style="margin-left: 2px;">Output File Name:</div>',
            unsafe_allow_html=True,
        )

        col_input, col_btn = st.columns([3.5, 1.2])

        with col_input:
            output_file_name_input = st.text_input(
                "Output File Name",
                value=default_out,
                key="filename_tab1",
                label_visibility="collapsed",
            )

        with col_btn:
            btn_click = st.button(
                "Format Report →",
                key="btn_gen_tab1",
                disabled=(uploaded_file is None),
            )

    if uploaded_file and btn_click:
        st.session_state["t1_processed"] = True
        with st.spinner("Formatting workbook..."):
            file_bytes = uploaded_file.read()
            df_raw = pd.read_excel(io.BytesIO(file_bytes))
            st.session_state["t1_buffer"] = process_rendered_medicines(df_raw)

    if st.session_state.get("t1_processed") and st.session_state.get("t1_buffer"):
        out_name = clean_filename(output_file_name_input, default_out)
        st.markdown(
            f"""
        <div class="success-banner">
            ✓ <strong>{out_name}</strong> processed successfully!
        </div>
        """,
            unsafe_allow_html=True,
        )

        st.download_button(
            label="📥 Download Formatted Excel Report",
            data=st.session_state["t1_buffer"],
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_tab1",
        )