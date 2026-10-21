# activity_logger.py
"""
Appends authentication and search events to the 'PortalActivity' Google Sheet.

Log tab layout (row 1 = headers, row 2+ = data):
    A = Serial (auto-increments from the highest existing value)
    B = Email
    C = Timestamp
    D = Activity ('Authenticated' / 'Authentication Failed' / roll number)
"""

import json
from datetime import datetime

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_NAME = "PortalActivity"
LOG_TAB          = "PortalLog"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
HEADERS = ["Serial", "Email", "Timestamp", "Activity"]


@st.cache_resource(show_spinner=False)
def _get_log_sheet():
    try:
        if "gcp_service_account_json" in st.secrets:
            creds_dict = json.loads(st.secrets["gcp_service_account_json"])
        elif "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            pk = creds_dict.get("private_key", "")
            if "\\n" in pk:
                creds_dict["private_key"] = pk.replace("\\n", "\n")
        else:
            return None, "No Google credentials in Secrets."
    except Exception as e:
        return None, f"Cannot parse secrets: {e}"

    for field in ("client_email", "private_key", "token_uri"):
        if not creds_dict.get(field):
            return None, f"Secrets missing required field: {field}"

    try:
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    except Exception as e:
        return None, f"Credentials failed: {e}"

    try:
        client = gspread.authorize(creds)
        sh = client.open(SPREADSHEET_NAME)
    except gspread.SpreadsheetNotFound:
        return None, (
            f"Spreadsheet '{SPREADSHEET_NAME}' not found. Share it with "
            f"{creds_dict['client_email']} as Editor."
        )
    except Exception as e:
        return None, f"client.open failed: {e}"

    try:
        try:
            ws = sh.worksheet(LOG_TAB)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=LOG_TAB, rows=1000, cols=4)
        return ws, None
    except Exception as e:
        return None, f"Opening log tab failed: {e}"


def _ensure_headers(sheet) -> None:
    try:
        first = sheet.row_values(1)
        if not first or len(first) < len(HEADERS):
            sheet.update("A1:D1", [HEADERS])
    except Exception:
        pass


def _next_serial(sheet) -> int:
    try:
        col_a = sheet.col_values(1)
    except Exception:
        return 1
    nums = []
    for v in col_a:
        s = str(v).strip()
        if not s:
            continue
        try:
            nums.append(int(float(s)))
        except ValueError:
            continue
    return (max(nums) + 1) if nums else 1


def log_activity(email: str, activity) -> None:
    sheet, err = _get_log_sheet()
    if sheet is None:
        st.warning(f"⚠️ Log skipped: {err}")
        return
    try:
        _ensure_headers(sheet)
        serial = _next_serial(sheet)
        ts = datetime.now().strftime("%d %b %Y %H:%M:%S")
        val = activity if isinstance(activity, int) else str(activity)
        sheet.append_row([serial, email, ts, val], value_input_option="USER_ENTERED")
        st.toast(f"📝 Logged #{serial}: {activity}", icon="✅")
    except Exception as e:
        st.warning(f"⚠️ Append failed: {e}")


def log_auth_success(email: str) -> None:
    log_activity(email, "Authenticated")


def log_auth_failure(email: str) -> None:
    log_activity(email, "Authentication Failed")


def log_search(email: str, roll: int) -> None:
    log_activity(email, int(roll))
