# activity_logger.py
from datetime import datetime
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

SHEET_NAME = "PortalActivity"
WORKSHEET = "Sheet1"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
HEADERS = ["Email", "Timestamp", "Activity"]

@st.cache_resource(show_spinner=False)
def _get_sheet():
    """Return (worksheet, error). Surface the real reason on failure."""
    if "gcp_service_account" not in st.secrets:
        return None, "Secrets block [gcp_service_account] not found."
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
    except Exception as e:
        return None, f"Cannot read secrets: {e}"

    # Check required fields
    for field in ("client_email", "private_key", "token_uri"):
        if not creds_dict.get(field):
            return None, f"Secrets missing required field: {field}"

    # Fix private key newlines if pasted with escaped \n
    pk = creds_dict["private_key"]
    if "\\n" in pk:
        creds_dict["private_key"] = pk.replace("\\n", "\n")

    try:
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    except Exception as e:
        return None, f"Credentials failed: {e}"

    try:
        client = gspread.authorize(creds)
    except Exception as e:
        return None, f"gspread authorize failed: {e}"

    try:
        sh = client.open(SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        return None, (
            f"Sheet '{SHEET_NAME}' not found. Check name and share it with "
            f"{creds_dict['client_email']} as Editor."
        )
    except Exception as e:
        return None, f"client.open failed: {e}"

    try:
        ws = sh.worksheet(WORKSHEET)
    except gspread.WorksheetNotFound:
        ws = sh.sheet1
    except Exception as e:
        return None, f"Opening worksheet failed: {e}"
    return ws, None

def _ensure_headers(sheet):
    try:
        first = sheet.row_values(1)
        if not first or len(first) < len(HEADERS):
            sheet.update("A1:C1", [HEADERS])
    except Exception:
        pass

def log_activity(email: str, activity):
    sheet, err = _get_sheet()
    if sheet is None:
        st.warning(f"⚠️ Log skipped: {err}")
        return
    try:
        _ensure_headers(sheet)
        ts = datetime.now().strftime("%d %b %Y %H:%M:%S")
        val = activity if isinstance(activity, int) else str(activity)
        sheet.append_row([email, ts, val], value_input_option="USER_ENTERED")
        st.toast(f"📝 Logged: {activity}", icon="✅")
    except Exception as e:
        st.warning(f"⚠️ Append failed: {e}")

def log_auth_success(email: str): log_activity(email, "Authenticated")
def log_auth_failure(email: str): log_activity(email, "Authentication Failed")
def log_search(email: str, roll: int): log_activity(email, int(roll))
