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
import time

import gspread
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

# 0-based column indices, computed from CHECKLIST_HEADERS
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
        "WAYBILL": False,  # Default unchecked boolean for Google Sheets
    }
    return [mapping.get(h, "") for h in CHECKLIST_HEADERS]


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
                "index": 1
            }
        })
    return requests


def _waybill_row_highlight_request(sheet_id, n_data_rows, n_cols, waybill_col_idx):
    """Highlights the ENTIRE row (from column A to N) in soft green when WAYBILL is checked."""
    col_letter = chr(65 + waybill_col_idx)  # Column 13 -> 'N'
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
            "index": 0  # Priority index 0 so it highlights the whole row when checked
        }
    }


def _build_navigation_tab(spreadsheet, sources_info):
    """Creates a 'Navigation' dashboard tab with big clickable source buttons using batched requests."""
    nav_title = "📌 Navigation"
    
    try:
        nav_ws = spreadsheet.worksheet(nav_title)
        nav_ws.clear()
        
        spreadsheet.batch_update({
            "requests": [{"unmergeCells": {"range": {"sheetId": nav_ws.id}}}]
        })
    except gspread.WorksheetNotFound:
        nav_ws = spreadsheet.add_worksheet(
            title=nav_title, rows=max(50, len(sources_info) * 4 + 10), cols=10
        )

    spreadsheet.reorder_worksheets([nav_ws] + [w for w in spreadsheet.worksheets() if w.title != nav_title])
    time.sleep(0.6)

    start_row = 5
    button_requests = []
    merge_requests = []
    
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
        
        merge_requests.append({
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

    nav_ws.batch_update(update_data, value_input_option="USER_ENTERED")
    time.sleep(0.8)

    if merge_requests:
        spreadsheet.batch_update({"requests": merge_requests})
        time.sleep(0.8)

    nav_ws.format("B2", {"textFormat": {"bold": True, "fontSize": 16, "foregroundColor": {"red": 0.1, "green": 0.2, "blue": 0.4}}})
    nav_ws.format("B3", {"textFormat": {"italic": True, "fontSize": 11, "foregroundColor": {"red": 0.4, "green": 0.4, "blue": 0.4}}})

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
    sources_info = []

    # Process worksheets safely one-by-one with strict pacing delays to prevent API 429 rate limit errors
    for source_label in sorted(rows_by_source.keys(), key=lambda s: s.lower()):
        rows = rows_by_source[source_label]
        title = _sanitize_sheet_title(source_label)
        n_data_rows = len(rows)
        n_cols = len(CHECKLIST_HEADERS)

        try:
            ws = spreadsheet.worksheet(title)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(
                title=title, rows=max(n_data_rows + 1, 2), cols=n_cols
            )
            time.sleep(1.0)  # Pause safely after creating a new worksheet tab

        values = [CHECKLIST_HEADERS] + [_row_to_values(r) for r in rows]
        
        # Combine value writing, bold header, and freeze row into a single batched payload
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": [{"range": f"{title}!A1", "values": values}]
        }
        client.values_batch_update(spreadsheet.id, body)
        
        # Collect formatting requests specific to this worksheet tab
        tab_formatting_requests = [
            {
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
            },
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": ws.id,
                        "gridProperties": {"frozenRowCount": 1}
                    },
                    "fields": "gridProperties.frozenRowCount"
                }
            }
        ]

        for col_idx in CHECKBOX_COL_INDICES:
            tab_formatting_requests.append(_checkbox_request(ws.id, n_data_rows, col_idx))
        
        tab_formatting_requests.append(
            _dropdown_request(ws.id, n_data_rows, CONSULT_COL_INDEX, CONSULT_OPTIONS)
        )
        tab_formatting_requests.append(
            _waybill_row_highlight_request(ws.id, n_data_rows, n_cols, WAYBILL_COL_INDEX)
        )
        tab_formatting_requests.extend(
            _consult_color_formatting_requests(ws.id, n_data_rows, CONSULT_COL_INDEX)
        )

        # Apply formatting per sheet in batch with a safe 1.2-second pause to strictly respect rate limits
        if tab_formatting_requests:
            spreadsheet.batch_update({"requests": tab_formatting_requests})
            time.sleep(1.2)  

        written.append((title, n_data_rows))
        sources_info.append((title, n_data_rows, ws.id))

    # Build Navigation Landing Page with batched merges and pauses
    _build_navigation_tab(spreadsheet, sources_info)

    return spreadsheet.url, written