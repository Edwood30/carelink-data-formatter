"""
FEATURE 3: Patient Follow-Up Checklist.

Upload the same two files as the Coordinator merge (Rendered Medicines +
Registered Patients). Instead of a downloadable file, this pushes the
result straight into a Google Sheet — one worksheet tab per Patient
Source, all inside a single fixed spreadsheet, updated in place on every
run. Requires Google Sheets access to be configured (see README.md).

Column detection is reused from feature_formatter (for the Rendered
Medicines file) and feature_merger (for the Registered Patients lookup)
rather than re-implemented, so fuzzy-matching behavior stays identical
across every feature in the app.
"""
import streamlit as st

import google_sheets
from feature_formatter import _detect_columns, _get
from feature_merger import _build_patient_lookup
from utils import clean_pin, clean_str, format_contact_number, read_data_file, split_name_fallback


def _last_first_middle(cols, row):
    last_val = clean_str(_get(row, cols["last"])) if cols["last"] else ""
    first_val = clean_str(_get(row, cols["first"])) if cols["first"] else ""
    middle_val = clean_str(_get(row, cols["middle"])) if cols["middle"] else ""
    raw_name = clean_str(_get(row, cols["patient_name"])) if cols["patient_name"] else ""

    if last_val and first_val:
        return last_val, first_val, middle_val
    if raw_name:
        return split_name_fallback(raw_name)
    return last_val, first_val, middle_val


def _build_rendered_rows(df_med):
    """One row per PATIENT (not per medicine) — medicines they received
    are aggregated into a single comma-separated cell. Grouped by the
    detected Patient Source for each patient."""
    cols = _detect_columns(df_med)
    by_pin = {}
    order = []

    for _, row in df_med.iterrows():
        pin = clean_pin(_get(row, cols["pin"])) if cols["pin"] else ""
        last, first, middle = _last_first_middle(cols, row)
        key = pin or f"NO_PIN::{last.lower()}::{first.lower()}"

        if key not in by_pin:
            by_pin[key] = {
                "Last Name": last,
                "First Name": first,
                "Middle Name": middle,
                "Cellphone Number": format_contact_number(_get(row, cols["phone"])) if cols["phone"] else "",
                "Full Address": clean_str(_get(row, cols["address"])) if cols["address"] else "",
                "Patient Source": clean_str(_get(row, cols["source"])) if cols["source"] else "",
                "Notes": clean_str(_get(row, cols["notes"])) if cols["notes"] else "",
                "_medicines": [],
                "_pin": pin,
            }
            order.append(key)

        medicine = clean_str(_get(row, cols["medicine"])) if cols["medicine"] else ""
        if medicine and medicine not in by_pin[key]["_medicines"]:
            by_pin[key]["_medicines"].append(medicine)
        if not by_pin[key]["Patient Source"] and cols["source"]:
            by_pin[key]["Patient Source"] = clean_str(_get(row, cols["source"]))

    return by_pin, order, cols


def build_checklist_rows_by_source(df_med, df_patient):
    """
    Returns (rows_by_source, seen_pins, med_cols, patient_meta, total_count)
    where rows_by_source is {source_label: [row_dict, ...]}, sorted
    alphabetically by Last Name (then First Name) within each source.
    Patients with no detected source (including everyone still awaiting
    follow-up) are grouped under "Unspecified Source".
    """
    by_pin, order, med_cols = _build_rendered_rows(df_med)

    patient_lookup, patient_meta = _build_patient_lookup(df_patient)
    if patient_lookup is None:
        return None, None, med_cols, None, 0

    all_rows = []
    seen_pins = set()

    for key in order:
        info = by_pin[key]
        if info["_pin"]:
            seen_pins.add(info["_pin"])
        all_rows.append(
            {
                "Last Name": info["Last Name"],
                "First Name": info["First Name"],
                "Middle Name": info["Middle Name"],
                "Cellphone Number": info["Cellphone Number"],
                "Full Address": info["Full Address"],
                "Medicines rendered": ", ".join(info["_medicines"]),
                "Patient Source": info["Patient Source"],
                "Notes": info["Notes"],
            }
        )

    for pin, info in patient_lookup.items():
        if pin in seen_pins:
            continue
        all_rows.append(
            {
                "Last Name": info["Last Name"],
                "First Name": info["First Name"],
                "Middle Name": info["Middle Name"],
                "Cellphone Number": info["Contact Number"],
                "Full Address": info["Full Address"],
                "Medicines rendered": "",
                "Patient Source": "",
                "Notes": "",
            }
        )

    all_rows.sort(key=lambda r: (r["Last Name"].strip().lower(), r["First Name"].strip().lower()))

    rows_by_source = {}
    for row in all_rows:
        source_label = row["Patient Source"].strip() or "Unspecified Source"
        rows_by_source.setdefault(source_label, []).append(row)

    return rows_by_source, seen_pins, med_cols, patient_meta, len(all_rows)


def render_checklist_tab():
    """Draws the full 'CareLink Checklist Data' tab UI and wires up processing."""
    st.markdown(
        """
    <div class="section-heading">Patient Follow-Up Checklist</div>
    <div class="section-subtext">Upload Rendered Medicines and Registered Patients files. Pushes one checklist tab per Patient Source straight into your Google Sheet — covering every patient, already served or still awaiting follow-up.</div>
    """,
        unsafe_allow_html=True,
    )

    if not google_sheets.is_configured():
        st.markdown(
            '<div class="card-title" style="color:#94A3B8;">'
            "Google Sheets isn't connected yet. See README.md for the "
            "setup steps (Google Cloud service account + Streamlit secrets) "
            "— this tab won't work until that's done.</div>",
            unsafe_allow_html=True,
        )
        return

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(
            """
        <div class="uploader-card">
            <div class="uploader-title uploader-accent-blue">01 — RENDERED MEDICINES FILE</div>
            <div class="uploader-sub">Upload Rendered Medicines (.xlsx or .csv)</div>
        </div>
        """,
            unsafe_allow_html=True,
        )
        file_med = st.file_uploader(
            "Upload Rendered Medicines",
            type=["xlsx", "csv"],
            key="tab3_med",
            label_visibility="collapsed",
        )

    with col2:
        st.markdown(
            """
        <div class="uploader-card">
            <div class="uploader-title uploader-accent-green">02 — REGISTERED PATIENTS FILE</div>
            <div class="uploader-sub">Upload Registered Patients (.xlsx or .csv)</div>
        </div>
        """,
            unsafe_allow_html=True,
        )
        file_patient = st.file_uploader(
            "Upload Registered Patients",
            type=["xlsx", "csv"],
            key="tab3_pat",
            label_visibility="collapsed",
        )

    def _sig(f):
        return f"{getattr(f, 'file_id', None) or f.name}-{f.size}" if f else None

    file_signature = (_sig(file_med), _sig(file_patient))
    if st.session_state.get("t3_file_signature") != file_signature:
        st.session_state["t3_file_signature"] = file_signature
        st.session_state["t3_processed"] = False
        st.session_state["t3_result"] = None

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        '<div class="card-title" style="margin-left: 2px;">Google Sheet:</div>',
        unsafe_allow_html=True,
    )

    c_url, c_action = st.columns([3.2, 1.2])
    with c_url:
        sheet_url_input = st.text_input(
            "Google Sheet URL",
            value=google_sheets.default_spreadsheet_url(),
            key="t3_sheet_url",
            label_visibility="collapsed",
            placeholder="Paste the Google Sheet URL…",
        )

    with c_action:
        both_uploaded = (file_med is not None) and (file_patient is not None)
        has_url = bool((sheet_url_input or "").strip())
        btn_click = st.button(
            "Push to Google Sheets", key="btn_gen_tab3", disabled=not (both_uploaded and has_url)
        )

    if not both_uploaded or not has_url:
        missing = []
        if file_med is None:
            missing.append("Rendered Medicines")
        if file_patient is None:
            missing.append("Registered Patients")
        if not has_url:
            missing.append("a Google Sheet URL")
        st.markdown(
            f'<div class="card-title" style="color:#94A3B8;margin-top:-4px;">'
            f'Please provide {" and ".join(missing)} before generating.</div>',
            unsafe_allow_html=True,
        )

    if both_uploaded and has_url and btn_click:
        st.session_state["t3_processed"] = True
        try:
            with st.spinner("Combining both files and writing to Google Sheets..."):
                df_med = read_data_file(file_med)
                df_patient = read_data_file(file_patient)
                rows_by_source, seen_pins, med_cols, patient_meta, total = build_checklist_rows_by_source(
                    df_med, df_patient
                )

                if rows_by_source is None:
                    st.error(
                        "Could not find a 'Patient PIN' column in the "
                        f"Registered Patients file. Found columns: {list(df_patient.columns)}"
                    )
                    st.session_state["t3_result"] = None
                else:
                    sheet_url, written = google_sheets.push_checklist_by_source(
                        sheet_url_input, rows_by_source
                    )
                    st.session_state["t3_result"] = {
                        "sheet_url": sheet_url,
                        "written": written,
                        "total": total,
                        "med_cols": med_cols,
                        "patient_meta": patient_meta,
                    }
        except Exception as e:
            st.session_state["t3_result"] = None
            hint = ""
            email = google_sheets.service_account_email()
            if email:
                hint = (
                    f" If this is a permissions error, make sure the sheet is "
                    f"shared with `{email}` as an Editor."
                )
            st.error(
                "Something went wrong while writing to Google Sheets. "
                "Please check the sheet URL and try again." + hint
            )
            with st.expander("Technical details"):
                st.code(str(e))

    result = st.session_state.get("t3_result")
    if st.session_state.get("t3_processed") and result:
        tabs_summary = ", ".join(f"{title} ({n} rows)" for title, n in result["written"])
        st.markdown(
            f"""
        <div class="success-banner">
            Updated {len(result['written'])} tab{'s' if len(result['written']) != 1 else ''} across {result['total']} patients: {tabs_summary}.
        </div>
        """,
            unsafe_allow_html=True,
        )
        st.markdown(f"[Open the spreadsheet]({result['sheet_url']})")

        with st.expander("Detected column mapping (click to verify)"):
            med_cols = result["med_cols"]
            patient_meta = result["patient_meta"]
            st.markdown(
                f"""
**Rendered Medicines file**
- Patient PIN → {f"`{med_cols['pin']}`" if med_cols['pin'] else "not found — left blank"}
- Name → {f"`{med_cols['last']}` + `{med_cols['first']}`" if med_cols['last'] and med_cols['first'] else (f"split from `{med_cols['patient_name']}`" if med_cols['patient_name'] else "not found — left blank")}
- Contact Number → {f"`{med_cols['phone']}`" if med_cols['phone'] else "not found — left blank"}
- Full Address → {f"`{med_cols['address']}`" if med_cols['address'] else "not found — left blank"}
- Medicine → {f"`{med_cols['medicine']}`" if med_cols['medicine'] else "not found — left blank"}
- Patient Source → {f"`{med_cols['source']}`" if med_cols['source'] else "not found — left blank"}

**Registered Patients file**
- Patient PIN → `{patient_meta['pin_col']}`
- Name → {patient_meta['name_source']}
- Contact Number → {f"`{patient_meta['phone_col']}`" if patient_meta['phone_col'] else "not found — left blank"}
- Full Address → {f"`{patient_meta['address_col']}`" if patient_meta['address_col'] else "not found — left blank"}
                """
            )