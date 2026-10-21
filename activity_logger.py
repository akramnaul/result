# activity_logger.py
from datetime import datetime
import json

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

SHEET_NAME = "PortalActivity"
WORKSHEET  = "Sheet1"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
HEADERS = ["Serial", "Email", "Timestamp", "Activity"]


@st.cache_resource(show_spinner=False)
def _get_sheet():
    """Return (worksheet, error). Surface real reason on failure."""
    # Two supported secrets formats:
    #   (A) gcp_service_account_json  -> whole JSON as one string (recommended)
    #   (B) [gcp_service_account]     -> TOML table
    try:
        if "gcp_service_account_json" in st.secrets:
            creds_dict = json.loads(st.secrets["gcp_service_account_json"])
        elif "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            # Heal private_key if pasted as escaped \n
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
        sh = client.open(SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        return None, (
            f"Sheet '{SHEET_NAME}' not found. Share it with "
            f"{creds_dict['client_email']} as Editor."
        )
    except Exception as e:
        return None, f"client.open failed: {e}"

    try:
        try:
            ws = sh.worksheet(WORKSHEET)
        except gspread.WorksheetNotFound:
            ws = sh.sheet1
        return ws, None
    except Exception as e:
        return None, f"Opening worksheet failed: {e}"


def _ensure_headers(sheet) -> None:
    """Ensure row 1 has the four column headers."""
    try:
        first = sheet.row_values(1)
        if not first or len(first) < len(HEADERS):
            sheet.update("A1:D1", [HEADERS])
    except Exception:
        pass


def _next_serial(sheet) -> int:
    """
    Read column A, find the highest numeric value, return it + 1.
    Starts at 1 if column A has no numeric entries.
    """
    try:
        col_a = sheet.col_values(1)   # includes header row
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
    sheet, err = _get_sheet()
    if sheet is None:
        st.warning(f"⚠️ Log skipped: {err}")
        return

    try:
        _ensure_headers(sheet)

        serial = _next_serial(sheet)
        ts = datetime.now().strftime("%d %b %Y %H:%M:%S")
        val = activity if isinstance(activity, int) else str(activity)

        sheet.append_row(
            [serial, email, ts, val],
            value_input_option="USER_ENTERED",
        )
        st.toast(f"📝 Logged #{serial}: {activity}", icon="✅")
    except Exception as e:
        st.warning(f"⚠️ Append failed: {e}")


def log_auth_success(email: str) -> None:
    log_activity(email, "Authenticated")


def log_auth_failure(email: str) -> None:
    log_activity(email, "Authentication Failed")


def log_search(email: str, roll: int) -> None:
    log_activity(email, int(roll))
