"""
Google Sheets integration for the Patient Follow-Up Checklist.

Instead of generating a downloadable .xlsx, this writes each Patient
Source into its own worksheet (tab) inside one fixed Google Spreadsheet
— same URL every time, one tab per source, updated in place on every run.

Requires a Google Cloud service account configured in Streamlit secrets.
See README.md for the full setup walkthrough. This module does nothing
(and is_configured() returns False) until that's done.
"""
import re

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

CHECKLIST_HEADERS = [
    "1st Contact",
    "Consult",
    "2nd Contact",
    "Last Name",
    "First Name",
    "Middle Name",
    "Cellphone Number",
    "Full Address",
    "Medicines rendered",
    "Patient Source",
    "Notes",
    "Prescribed",
    "Packed",
]

# 0-based column indices, computed from CHECKLIST_HEADERS
_CHECKBOX_HEADERS = ("1st Contact", "2nd Contact", "Prescribed", "Packed")
CHECKBOX_COL_INDICES = [CHECKLIST_HEADERS.index(h) for h in _CHECKBOX_HEADERS]
CONSULT_COL_INDEX = CHECKLIST_HEADERS.index("Consult")

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
    """Accepts either a full Google Sheets URL or a bare spreadsheet ID."""
    url_or_id = (url_or_id or "").strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url_or_id)
    return m.group(1) if m else url_or_id


@st.cache_resource(show_spinner=False)
def _get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def _sanitize_sheet_title(s):
    """Google Sheets tab names can't contain : \\ / ? * [ ] and are capped at 100 chars."""
    s = str(s or "").strip()
    for ch in ("\\", "/", "?", "*", "[", "]", ":"):
        s = s.replace(ch, "")
    s = re.sub(r"\s+", " ", s).strip()
    return (s[:100] or "Unspecified Source")


def _row_to_values(row):
    return [
        "" if h in _CHECKBOX_HEADERS or h == "Consult" else row.get(h, "")
        for h in CHECKLIST_HEADERS
    ]


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
    """Adds conditional formatting rules to color the Consult status options."""
    color_map = {
        "Pending": {
            "bg": {"red": 1.0, "green": 0.95, "blue": 0.8},
            "fg": {"red": 0.5, "green": 0.35, "blue": 0.0},
        },
        "Scheduled": {
            "bg": {"red": 0.88, "green": 0.93, "blue": 1.0},
            "fg": {"red": 0.1, "green": 0.3, "blue": 0.6},
        },
        "Completed": {
            "bg": {"red": 0.85, "green": 0.94, "blue": 0.85},
            "fg": {"red": 0.1, "green": 0.4, "blue": 0.1},
        },
        "No Show": {
            "bg": {"red": 0.98, "green": 0.85, "blue": 0.85},
            "fg": {"red": 0.6, "green": 0.1, "blue": 0.1},
        },
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
                "index": 0
            }
        })
    return requests


def _build_navigation_tab(spreadsheet, sources_info):
    """Creates a 'Navigation' dashboard tab with big clickable source buttons."""
    nav_title = "📌 Navigation"
    
    try:
        nav_ws = spreadsheet.worksheet(nav_title)
        nav_ws.clear()
    except gspread.WorksheetNotFound:
        nav_ws = spreadsheet.add_worksheet(title=nav_title, rows=50, cols=10)

    # Reorder Navigation sheet to be first (index 0)
    spreadsheet.reorder_worksheets([nav_ws] + [w for w in spreadsheet.worksheets() if w.title != nav_title])

    # Setup Title Banner
    values = [["CareLink Patient Checklist Navigation"], ["Select a Patient Source below to jump to its sheet:"], []]
    
    start_row = 5
    button_requests = []

    # Format Title
    nav_ws.update(values=[["📌 CareLink Patient Checklist Navigation"], ["Click any button below to jump directly to that source sheet:"]], range_name="B2:B3")
    nav_ws.format("B2", {"textFormat": {"bold": True, "fontSize": 16, "foregroundColor": {"red": 0.1, "green": 0.2, "blue": 0.4}}})
    nav_ws.format("B3", {"textFormat": {"italic": True, "fontSize": 11, "foregroundColor": {"red": 0.4, "green": 0.4, "blue": 0.4}}})

    # Create Big Button Rows
    for idx, (title, count, gid) in enumerate(sources_info):
        row_num = start_row + (idx * 3)  # 3 rows spacing per button
        cell_range = f"B{row_num}:E{row_num+1}" # Merged 2x4 block for BIG button appearance
        
        formula = f'=HYPERLINK("#gid={gid}", "📂 {title.upper()} ({count} Patients)")'
        
        nav_ws.update(values=[[formula]], range_name=f"B{row_num}")
        nav_ws.merge_cells(cell_range)
        
        # Style as a big button
        button_requests.append({
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
                        "backgroundColor": {"red": 0.15, "green": 0.45, "blue": 0.85}, # Soft Blue Button
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {
                            "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, # White text
                            "bold": True,
                            "fontSize": 13,
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)",
            }
        })

    if button_requests:
        spreadsheet.batch_update({"requests": button_requests})


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
    sources_info = [] # Stores (title, count, sheet_id)
    formatting_requests = []

    for source_label in sorted(rows_by_source.keys(), key=lambda s: s.lower()):
        rows = rows_by_source[source_label]
        title = _sanitize_sheet_title(source_label)
        n_data_rows = len(rows)
        n_cols = len(CHECKLIST_HEADERS)

        try:
            ws = spreadsheet.worksheet(title)
            ws.clear()
            ws.resize(rows=max(n_data_rows + 1, 2), cols=n_cols)
        except gspread.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(
                title=title, rows=max(n_data_rows + 1, 2), cols=n_cols
            )

        values = [CHECKLIST_HEADERS] + [_row_to_values(r) for r in rows]
        ws.update(values=values, range_name="A1")
        ws.format("1:1", {"textFormat": {"bold": True}})
        ws.freeze(rows=1)

        # Checkbox data validation
        for col_idx in CHECKBOX_COL_INDICES:
            formatting_requests.append(_checkbox_request(ws.id, n_data_rows, col_idx))
        
        # Consult dropdown validation
        formatting_requests.append(
            _dropdown_request(ws.id, n_data_rows, CONSULT_COL_INDEX, CONSULT_OPTIONS)
        )

        # Consult option colors (Pending, Scheduled, Completed, No Show)
        formatting_requests.extend(
            _consult_color_formatting_requests(ws.id, n_data_rows, CONSULT_COL_INDEX)
        )

        written.append((title, n_data_rows))
        sources_info.append((title, n_data_rows, ws.id))

    if formatting_requests:
        spreadsheet.batch_update({"requests": formatting_requests})

    # Build Navigation Landing Page with Big Buttons
    _build_navigation_tab(spreadsheet, sources_info)

    return spreadsheet.url, written