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

# 0-based column indices, computed from CHECKLIST_HEADERS so they can
# never silently drift out of sync if the header list is edited later.
_CHECKBOX_HEADERS = ("1st Contact", "2nd Contact", "Prescribed", "Packed")
CHECKBOX_COL_INDICES = [CHECKLIST_HEADERS.index(h) for h in _CHECKBOX_HEADERS]
CONSULT_COL_INDEX = CHECKLIST_HEADERS.index("Consult")

# Assumption (no explicit values given): a reasonable default follow-up
# vocabulary. Change freely.
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
    """Google Sheets tab names can't contain : \\ / ? * [ ] and are capped
    at 100 characters."""
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


def _text_wrap_request(sheet_id, n_data_rows, n_cols):
    """Enables WRAP strategy for all cells in the worksheet."""
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 1 + n_data_rows,
                "startColumnIndex": 0,
                "endColumnIndex": n_cols,
            },
            "cell": {
                "userEnteredFormat": {
                    "wrapStrategy": "WRAP"
                }
            },
            "fields": "userEnteredFormat.wrapStrategy",
        }
    }


def _consult_color_formatting_requests(sheet_id, n_data_rows, consult_col_idx):
    """Adds conditional formatting rules to color the Consult status options."""
    color_map = {
        "Pending": {
            "bg": {"red": 1.0, "green": 0.95, "blue": 0.8},      # Soft Yellow
            "fg": {"red": 0.5, "green": 0.35, "blue": 0.0},
        },
        "Scheduled": {
            "bg": {"red": 0.88, "green": 0.93, "blue": 1.0},     # Soft Blue
            "fg": {"red": 0.1, "green": 0.3, "blue": 0.6},
        },
        "Completed": {
            "bg": {"red": 0.85, "green": 0.94, "blue": 0.85},    # Soft Green
            "fg": {"red": 0.1, "green": 0.4, "blue": 0.1},
        },
        "No Show": {
            "bg": {"red": 0.98, "green": 0.85, "blue": 0.85},     # Soft Red
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


def push_checklist_by_source(spreadsheet_url_or_id, rows_by_source):
    client = _get_client()

    sheet_id = extract_spreadsheet_id(spreadsheet_url_or_id)

    if not sheet_id:
        raise RuntimeError(
            "No Google Spreadsheet ID was found."
        )

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

        # Cell Text Wrapping
        formatting_requests.append(_text_wrap_request(ws.id, n_data_rows, n_cols))

        # Consult option colors (Pending, Scheduled, Completed, No Show)
        formatting_requests.extend(
            _consult_color_formatting_requests(ws.id, n_data_rows, CONSULT_COL_INDEX)
        )

        written.append((title, n_data_rows))

    if formatting_requests:
        spreadsheet.batch_update({"requests": formatting_requests})

    return spreadsheet.url, written