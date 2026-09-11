"""
Google Sheets integration for the Patient Follow-Up Checklist.

Instead of generating a downloadable .xlsx, this writes each Patient
Source into its own worksheet (tab) inside one fixed Google Spreadsheet
— same URL every time, updated in place on every run.

Requires a Google Cloud service account configured in Streamlit secrets.
"""
import re
import time

import gspread
from gspread.exceptions import APIError
import streamlit as st
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Preset Google Sheets mappings for the dropdown selector
PRESET_SPREADSHEETS = {
    "NCJJ SOURCES": "https://docs.google.com/spreadsheets/d/1hGFuP1OiZxT2XjXI_RE-7UQ-C9tEk9dV_cJc1u4c7Gg/edit?gid=861907105#gid=861907105",
    "CALOOCAN SITES": "https://docs.google.com/spreadsheets/d/1xDj3gy4ha-D9yXext4vVQeaECs-QdirIAqWcSp5dU_M/edit?gid=22834876#gid=22834876",
    "PRIVATE COMPANIES": "https://docs.google.com/spreadsheets/d/1y92gEzc2D8kHeO4r4lE9_Hd4SlURt_WYkPqUqKGs-GE/edit?gid=492635901#gid=492635901",
}

CHECKLIST_HEADERS = [
    "PATIENT SOURCE",
    "CONTACT NUMBER",
    "LAST NAME",
    "FIRST NAME",
    "MIDDLE NAME",
    "ADDRESS",
    "1ST CONTACT",
    "NOTES",
    "CONSULT",
    "MEDICINE RENDERED",
    "PRESCRIBED",
    "PACKED",
    "2ND CONTACT",
    "WAYBILL",
]

_CHECKBOX_HEADERS = ("1ST CONTACT", "2ND CONTACT", "PRESCRIBED", "PACKED", "WAYBILL")
CHECKBOX_COL_INDICES = [CHECKLIST_HEADERS.index(h) for h in _CHECKBOX_HEADERS]
CONSULT_COL_INDEX = CHECKLIST_HEADERS.index("CONSULT")
WAYBILL_COL_INDEX = CHECKLIST_HEADERS.index("WAYBILL")

CONSULT_OPTIONS = ["Pending", "Scheduled", "Completed", "No Show"]


def is_configured():
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


def service_account_email():
    try:
        return st.secrets["gcp_service_account"]["client_email"]
    except Exception:
        return None


def default_spreadsheet_url():
    try:
        return st.secrets.get("default_spreadsheet_url", "")
    except Exception:
        return ""


def extract_spreadsheet_id(url_or_id):
    url_or_id = (url_or_id or "").strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url_or_id)
    return m.group(1) if m else url_or_id


@st.cache_resource(show_spinner=False)
def _get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def _sanitize_sheet_title(s):
    s = str(s or "").strip()
    for ch in ("\\", "/", "?", "*", "[", "]", ":"):
        s = s.replace(ch, "")
    s = re.sub(r"\s+", " ", s).strip()
    return (s[:100] or "Unspecified Source")


def _row_to_values(row):
    mapping = {
        "PATIENT SOURCE": row.get("PATIENT SOURCE", row.get("Patient Source", "")),
        "CONTACT NUMBER": row.get("CONTACT NUMBER", row.get("Cellphone Number", "")),
        "LAST NAME": row.get("LAST NAME", row.get("Last Name", "")),
        "FIRST NAME": row.get("FIRST NAME", row.get("First Name", "")),
        "MIDDLE NAME": row.get("MIDDLE NAME", row.get("Middle Name", "")),
        "ADDRESS": row.get("ADDRESS", row.get("Full Address", "")),
        "1ST CONTACT": "",
        "NOTES": row.get("NOTES", row.get("Notes", "")),
        "CONSULT": "",
        "MEDICINE RENDERED": row.get("MEDICINE RENDERED", row.get("Medicines rendered", "")),
        "PRESCRIBED": "",
        "PACKED": "",
        "2ND CONTACT": "",
        "WAYBILL": False, 
    }
    return [mapping.get(h, "") for h in CHECKLIST_HEADERS]


def _safe_batch_update(func, *args, **kwargs):
    """Executes a batch update with exponential backoff to handle 429 Quota Exceeded errors."""
    max_retries = 4
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except APIError as e:
            if "429" in str(e) and attempt < max_retries - 1:
                time.sleep((2 ** attempt) + 1) # Sleeps for 2, 5, 9 seconds...
            else:
                raise


def _checkbox_request(sheet_id, n_data_rows, col_index):
    return {
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,
                "endRowIndex": 1 + n_data_rows,
                "startColumnIndex": col_index,
                "endColumnIndex": col_index + 1,
            },
            "rule": {"condition": {"type": "BOOLEAN"}, "strict": True},
        }
    }


def _dropdown_request(sheet_id, n_data_rows, col_index, options):
    return {
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,
                "endRowIndex": 1 + n_data_rows,
                "startColumnIndex": col_index,
                "endColumnIndex": col_index + 1,
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": o} for o in options],
                },
                "showCustomUi": True,
                "strict": True,
            },
        }
    }


def _consult_color_formatting_requests(sheet_id, n_data_rows, consult_col_idx):
    color_map = {
        "Pending": {"bg": {"red": 1.0, "green": 0.95, "blue": 0.8}, "fg": {"red": 0.5, "green": 0.35, "blue": 0.0}},
        "Scheduled": {"bg": {"red": 0.88, "green": 0.93, "blue": 1.0}, "fg": {"red": 0.1, "green": 0.3, "blue": 0.6}},
        "Completed": {"bg": {"red": 0.85, "green": 0.94, "blue": 0.85}, "fg": {"red": 0.1, "green": 0.4, "blue": 0.1}},
        "No Show": {"bg": {"red": 0.98, "green": 0.85, "blue": 0.85}, "fg": {"red": 0.6, "green": 0.1, "blue": 0.1}},
    }

    requests = []
    for option, colors in color_map.items():
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": 1 + n_data_rows,
                        "startColumnIndex": consult_col_idx,
                        "endColumnIndex": consult_col_idx + 1,
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "TEXT_EQ",
                            "values": [{"userEnteredValue": option}]
                        },
                        "format": {
                            "backgroundColor": colors["bg"],
                            "textFormat": {"foregroundColor": colors["fg"], "bold": True}
                        }
                    }
                },
                "index": 1
            }
        })
    return requests


def _waybill_row_highlight_request(sheet_id, n_data_rows, n_cols, waybill_col_idx):
    col_letter = chr(65 + waybill_col_idx) 
    formula = f"=${col_letter}2=TRUE"

    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 1 + n_data_rows,
                    "startColumnIndex": 0,
                    "endColumnIndex": n_cols,
                }],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": formula}]
                    },
                    "format": {
                        "backgroundColor": {"red": 0.8, "green": 0.93, "blue": 0.8},
                        "textFormat": {"foregroundColor": {"red": 0.05, "green": 0.35, "blue": 0.05}}
                    }
                }
            },
            "index": 0 
        }
    }


def _build_navigation_tab(spreadsheet, sources_info):
    """Creates a 'Navigation' dashboard using exactly TWO quota-efficient batch calls."""
    nav_title = "📌 Navigation"
    
    try:
        nav_ws = spreadsheet.worksheet(nav_title)
        nav_ws.clear()
    except gspread.WorksheetNotFound:
        nav_ws = spreadsheet.add_worksheet(title=nav_title, rows=max(50, len(sources_info) * 4 + 10), cols=10)

    spreadsheet.reorder_worksheets([nav_ws] + [w for w in spreadsheet.worksheets() if w.title != nav_title])
    
    # Payload 1: Update all text values
    start_row = 5
    update_data = [
        {
            "range": "B2:B3",
            "values": [
                ["📌 CareLink Patient Checklist Navigation"],
                ["Click any button below to jump directly to that source sheet:"]
            ]
        }
    ]

    for idx, (title, count, gid) in enumerate(sources_info):
        row_num = start_row + (idx * 3)
        formula = f'=HYPERLINK("#gid={gid}", "📂 {title.upper()} ({count} Patients)")'
        update_data.append({
            "range": f"B{row_num}",
            "values": [[formula]]
        })

    _safe_batch_update(nav_ws.batch_update, update_data, value_input_option="USER_ENTERED")

    # Payload 2: Compile ALL formatting rules, unmerges, and styles into ONE request
    formatting_requests = [{"unmergeCells": {"range": {"sheetId": nav_ws.id}}}]
    
    # Title Formats B2 & B3 (Replaces the unbatched nav_ws.format calls)
    formatting_requests.append({
        "repeatCell": {
            "range": {"sheetId": nav_ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 1, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 16, "foregroundColor": {"red": 0.1, "green": 0.2, "blue": 0.4}}}},
            "fields": "userEnteredFormat.textFormat"
        }
    })
    formatting_requests.append({
        "repeatCell": {
            "range": {"sheetId": nav_ws.id, "startRowIndex": 2, "endRowIndex": 3, "startColumnIndex": 1, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": {"textFormat": {"italic": True, "fontSize": 11, "foregroundColor": {"red": 0.4, "green": 0.4, "blue": 0.4}}}},
            "fields": "userEnteredFormat.textFormat"
        }
    })

    for idx, (title, count, gid) in enumerate(sources_info):
        row_num = start_row + (idx * 3)
        
        formatting_requests.append({
            "mergeCells": {
                "range": {
                    "sheetId": nav_ws.id,
                    "startRowIndex": row_num - 1,
                    "endRowIndex": row_num + 1,
                    "startColumnIndex": 1,
                    "endColumnIndex": 5,
                },
                "mergeType": "MERGE_ALL"
            }
        })
        
        formatting_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": nav_ws.id,
                    "startRowIndex": row_num - 1,
                    "endRowIndex": row_num + 1,
                    "startColumnIndex": 1,
                    "endColumnIndex": 5,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": 0.15, "green": 0.45, "blue": 0.85},
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {
                            "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                            "bold": True,
                            "fontSize": 13,
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)",
            }
        })

    if formatting_requests:
        _safe_batch_update(spreadsheet.batch_update, {"requests": formatting_requests})


def push_checklist_by_source(spreadsheet_url_or_id, rows_by_source):
    client = _get_client()
    sheet_id = extract_spreadsheet_id(spreadsheet_url_or_id)

    if not sheet_id:
        raise RuntimeError("No Google Spreadsheet ID was found.")

    try:
        spreadsheet = client.open_by_key(sheet_id)
    except Exception as e:
        raise RuntimeError(
            f"Could not open spreadsheet.\n"
            f"Spreadsheet ID: {sheet_id}\n"
            f"Exception: {type(e).__name__}\n"
            f"Details: {repr(e)}"
        ) from e
    
    written = []
    sources_info = []

    existing_worksheets = {ws.title: ws for ws in spreadsheet.worksheets()}

    missing_titles = [
        _sanitize_sheet_title(label) 
        for label in rows_by_source.keys() 
        if _sanitize_sheet_title(label) not in existing_worksheets and _sanitize_sheet_title(label) != "📌 Navigation"
    ]
    
    if missing_titles:
        add_requests = [{"addSheet": {"properties": {"title": title}}} for title in missing_titles]
        _safe_batch_update(spreadsheet.batch_update, {"requests": add_requests})
        existing_worksheets = {ws.title: ws for ws in spreadsheet.worksheets()}

    batch_values_data = []
    all_formatting_requests = []

    sorted_sources = sorted(rows_by_source.keys(), key=lambda s: s.lower())

    for source_label in sorted_sources:
        title = _sanitize_sheet_title(source_label)
        rows = rows_by_source[source_label]
        n_data_rows = len(rows)
        n_cols = len(CHECKLIST_HEADERS)
        ws = existing_worksheets.get(title)

        if title not in missing_titles:
            if ws:
                sources_info.append((title, f"Existing (New: {n_data_rows})", ws.id))
            continue

        values = [CHECKLIST_HEADERS] + [_row_to_values(r) for r in rows]
        batch_values_data.append({
            "range": f"'{title}'!A1",
            "values": values
        })

        all_formatting_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": n_cols
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"bold": True}
                    }
                },
                "fields": "userEnteredFormat.textFormat.bold"
            }
        })
        all_formatting_requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": ws.id,
                    "gridProperties": {"frozenRowCount": 1}
                },
                "fields": "gridProperties.frozenRowCount"
            }
        })

        for col_idx in CHECKBOX_COL_INDICES:
            all_formatting_requests.append(_checkbox_request(ws.id, n_data_rows, col_idx))
        
        all_formatting_requests.append(
            _dropdown_request(ws.id, n_data_rows, CONSULT_COL_INDEX, CONSULT_OPTIONS)
        )
        all_formatting_requests.append(
            _waybill_row_highlight_request(ws.id, n_data_rows, n_cols, WAYBILL_COL_INDEX)
        )
        all_formatting_requests.extend(
            _consult_color_formatting_requests(ws.id, n_data_rows, CONSULT_COL_INDEX)
        )

        written.append((title, n_data_rows))
        sources_info.append((title, n_data_rows, ws.id))

    if batch_values_data:
        _safe_batch_update(spreadsheet.values_batch_update, {
            "valueInputOption": "USER_ENTERED",
            "data": batch_values_data
        })

    if all_formatting_requests:
        chunk_size = 500
        for i in range(0, len(all_formatting_requests), chunk_size):
            chunk = all_formatting_requests[i:i + chunk_size]
            _safe_batch_update(spreadsheet.batch_update, {"requests": chunk})

    _build_navigation_tab(spreadsheet, sources_info)

    return spreadsheet.url, written