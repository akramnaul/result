# activity_logger.py
"""
Activity logger for the FCC Result Portal.

Writes authentication and search events to a Google Sheet named
'PortalActivity'. Fails gracefully (visible warning, no crash) if the
Google credentials are missing or the sheet is unreachable.
"""

from datetime import datetime

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ------------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------------
SHEET_NAME = "PortalActivity"          # Google Sheets file name
WORKSHEET  = "Sheet1"                  # tab name inside that file
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

HEADERS = ["Email", "Timestamp", "Activity"]


# ------------------------------------------------------------------
# Internal: cached gspread Worksheet
# ------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def _get_sheet():
    """
    Return the gspread Worksheet, or None if credentials are absent
    or authentication fails. Cached once per Streamlit session.
    """
    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=SCOPES,
        )
        client = gspread.authorize(creds)
        # Try the named worksheet first, fall back to the first sheet
        try:
            return client.open(SHEET_NAME).worksheet(WORKSHEET)
        except gspread.WorksheetNotFound:
            return client.open(SHEET_NAME).sheet1
    except KeyError:
        return None          # [gcp_service_account] block missing from secrets
    except Exception:
        return None


def _ensure_headers(sheet) -> None:
    """Ensure row 1 contains the expected column headers."""
    try:
        first_row = sheet.row_values(1)
        if not first_row or len(first_row) < len(HEADERS):
            sheet.update("A1:C1", [HEADERS])
    except Exception:
        pass


# ------------------------------------------------------------------
# Public API
# ------------------------------------------------------------------
def log_activity(email: str, activity) -> None:
    """
    Append one row to the PortalActivity Google Sheet.

    Args:
        email:    faculty email (string)
        activity: 'Authenticated' / 'Authentication Failed' (string),
                  or a roll number (int) for search events.

    Never raises. Failures are surfaced via st.warning().
    """
    sheet = _get_sheet()
    if sheet is None:
        return

    try:
        _ensure_headers(sheet)

        timestamp = datetime.now().strftime("%d %b %Y %H:%M:%S")

        if isinstance(activity, int) and not isinstance(activity, bool):
            activity_cell = activity
        else:
            activity_cell = str(activity)

        sheet.append_row(
            [email, timestamp, activity_cell],
            value_input_option="USER_ENTERED",
        )
    except Exception as e:
        st.warning(f"⚠️ Log not saved to Google Sheets: {type(e).__name__}: {e}")


def log_auth_success(email: str) -> None:
    log_activity(email, "Authenticated")


def log_auth_failure(email: str) -> None:
    log_activity(email, "Authentication Failed")


def log_search(email: str, roll: int) -> None:
    log_activity(email, int(roll))
