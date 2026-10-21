# streamlit_app.py
"""
FCC Result Portal — reads data from the 'PortalActivity' Google Sheet
(tabs: FacultyStaffList, 26-09, 26-10, …) and writes activity logs
back into the same spreadsheet.

Result tab layout:
    A = Roll Number
    B = Overall Attendance (%)
    C = Student Name
    D..M = 10 Subject marks (%)
"""

import json
import re

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

from activity_logger import (
    log_auth_success,
    log_auth_failure,
    log_search,
)

# ============================================================
# CONFIG
# ============================================================
SPREADSHEET_NAME = "PortalActivity"

AUTH_TAB = "FacultyStaffList"
AUTH_EMAIL_COLUMN = "G"

RESULT_TABS = ["26-09", "26-10"]
TAB_LABELS = {
    "26-09": "September 2026",
    "26-10": "October 2026",
}

# Column layout inside each result tab (Excel-style letters)
ROLL_COL       = "A"    # Roll number
ATTENDANCE_COL = "B"    # Overall attendance %
NAME_COL       = "C"    # Student name
SUBJECT_COLS   = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]

# Column indices (0-based)
IDX_ROLL       = col_letter_to_index if False else 0   # A
IDX_ATTENDANCE = 1                                     # B
IDX_NAME       = 2                                     # C
IDX_SUBJECT_START = 3                                  # D

MAX_SECTION_LOOKBACK     = 50
MIN_HEADER_SUBJECT_CELLS = 3

CACHE_TTL_SECONDS = 60

_INVISIBLE = (
    "\u200b", "\u200c", "\u200d", "\u2060", "\ufeff",
    "\xa0", "\u2009", "\u202f", "\u00ad", "\u200e", "\u200f",
    "\u202a", "\u202b", "\u202c", "\u202d", "\u202e",
)

# ============================================================
# HELPERS
# ============================================================
def col_letter_to_index(letter: str) -> int:
    letter = letter.upper()
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def normalize(value) -> str:
    if value is None: return ""
    if isinstance(value, float) and pd.isna(value): return ""
    s = str(value)
    for ch in _INVISIBLE: s = s.replace(ch, "")
    s = s.strip().strip("'\"`‘’“”").strip()
    return s.lower()


def normalize_roll(value):
    if value is None: return None
    if isinstance(value, bool): return None
    if isinstance(value, float) and pd.isna(value): return None
    if isinstance(value, int): return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else None
    s = str(value)
    for ch in _INVISIBLE: s = s.replace(ch, "")
    s = s.strip().strip("'\"`‘’“”").strip()
    if not s: return None
    s_clean = s.replace(",", "").replace(" ", "")
    try: return int(s_clean)
    except ValueError: pass
    try:
        f = float(s_clean)
        return int(f) if f.is_integer() else None
    except ValueError: pass
    digits = re.sub(r"\D", "", s)
    return int(digits) if digits else None


_PCT_RE = re.compile(r"-?\d+(?:\.\d+)?")


def format_percent(raw) -> str:
    if raw is None: return "—"
    if isinstance(raw, float) and pd.isna(raw): return "—"
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        val = float(raw)
        if val == 1.0: return "100%"
        if 0.0 < val < 1.0: return f"{round(val * 100)}%"
        return f"{round(val)}%"
    s = str(raw)
    for ch in _INVISIBLE: s = s.replace(ch, "")
    s = s.strip()
    if not s: return "—"
    if "%" in s:
        m = _PCT_RE.search(s.replace(",", ""))
        if not m: return "—"
        try: return f"{round(float(m.group(0)))}%"
        except ValueError: return "—"
    try: val = float(s.replace(",", ""))
    except ValueError: return s
    if val == 1.0: return "100%"
    if 0.0 < val < 1.0: return f"{round(val * 100)}%"
    return f"{round(val)}%"


# ============================================================
# GOOGLE SHEETS CLIENT
# ============================================================
@st.cache_resource(show_spinner=False)
def _gsheet_client():
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
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        return gspread.authorize(creds), None
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"


@st.cache_resource(show_spinner=False)
def _open_spreadsheet():
    client, err = _gsheet_client()
    if client is None:
        return None, err
    try:
        return client.open(SPREADSHEET_NAME), None
    except gspread.SpreadsheetNotFound:
        return None, (
            f"Spreadsheet '{SPREADSHEET_NAME}' not found. Share it with the "
            f"service account as Editor."
        )
    except Exception as e:
        return None, f"client.open failed: {type(e).__name__}: {e}"


@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def load_tab_as_df(tab_name: str):
    sh, _ = _open_spreadsheet()
    if sh is None:
        return None
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        return None
    except Exception:
        return None
    try:
        values = ws.get_all_values()
    except Exception:
        return None
    if not values:
        return pd.DataFrame()
    header, *rows = values
    width = len(header)
    norm_rows = [r + [""] * (width - len(r)) for r in rows]
    df = pd.DataFrame(norm_rows, columns=header)
    df.columns = [str(c).strip() for c in df.columns]
    return df.replace("", pd.NA)


@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def available_tabs() -> list[str]:
    sh, _ = _open_spreadsheet()
    if sh is None:
        return []
    try:
        return [w.title for w in sh.worksheets()]
    except Exception:
        return []


# ============================================================
# AUTHENTICATION
# ============================================================
@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def accepted_emails() -> set[str]:
    df = load_tab_as_df(AUTH_TAB)
    if df is None or df.empty:
        return set()
    emails: set[str] = set()

    g_idx = col_letter_to_index(AUTH_EMAIL_COLUMN)
    if g_idx < len(df.columns):
        for v in df.iloc[:, g_idx].dropna():
            n = normalize(v)
            if "@" in n:
                emails.add(n)

    for j in range(df.shape[1]):
        col_vals = df.iloc[:, j].dropna().astype(str)
        if len(col_vals) == 0:
            continue
        hits = sum(1 for v in col_vals if "@" in v)
        if hits >= 3 and hits >= len(col_vals) * 0.3:
            for v in col_vals:
                n = normalize(v)
                if "@" in n:
                    emails.add(n)
    return emails


# ============================================================
# ROLL LOOKUP
# ============================================================
def find_roll_row(df: pd.DataFrame, roll: int):
    target = int(roll)
    for j in range(df.shape[1]):
        parsed = df.iloc[:, j].map(normalize_roll)
        hits = parsed[parsed == target]
        if len(hits):
            return hits.index[0]
    return None


# ============================================================
# SUBJECT NAME (section header)
# ============================================================
def find_section_header_row(df, student_row, roll_col_idx=IDX_ROLL):
    if roll_col_idx >= len(df.columns):
        return None
    subject_idx = [j for j in range(len(df.columns)) if j >= IDX_SUBJECT_START]
    start = student_row - 1
    stop = max(0, student_row - MAX_SECTION_LOOKBACK)
    for r in range(start, stop - 1, -1):
        v = df.iloc[r, roll_col_idx]
        if pd.isna(v) or str(v).strip() == "":
            filled = sum(
                1 for j in subject_idx
                if not pd.isna(df.iloc[r, j]) and str(df.iloc[r, j]).strip() != ""
            )
            if filled >= MIN_HEADER_SUBJECT_CELLS:
                return r
    return None


def subject_name_for_section(df, student_row, subject_col_idx) -> str:
    if subject_col_idx >= len(df.columns):
        return "NotUploaded"
    hdr = find_section_header_row(df, student_row, IDX_ROLL)
    if hdr is None:
        return "NotUploaded"
    v = df.iloc[hdr, subject_col_idx]
    if pd.isna(v):
        return "NotUploaded"
    s = str(v).strip()
    if not s:
        return "NotUploaded"
    try:
        float(s.replace(",", "").replace("%", "").strip())
        return "NotUploaded"
    except ValueError:
        return s


# ============================================================
# RECORD EXTRACTION
# ============================================================
def safe_str_cell(df, row_idx, col_idx):
    if col_idx >= len(df.columns):
        return ""
    v = df.iloc[row_idx, col_idx]
    return "" if pd.isna(v) else str(v).strip()


def extract_record(df: pd.DataFrame, row_idx: int) -> dict:
    raw_roll = df.iloc[row_idx, IDX_ROLL] if df.shape[1] > IDX_ROLL else None
    roll_int = normalize_roll(raw_roll)
    roll_display = (
        str(roll_int) if roll_int is not None
        else safe_str_cell(df, row_idx, IDX_ROLL)
    )

    # Attendance: column B
    attendance = format_percent(
        df.iloc[row_idx, IDX_ATTENDANCE]
        if df.shape[1] > IDX_ATTENDANCE and not pd.isna(df.iloc[row_idx, IDX_ATTENDANCE])
        else None
    )

    # Name: column C
    name = safe_str_cell(df, row_idx, IDX_NAME)

    # Subjects: columns D..M
    subjects = []
    for j in range(IDX_SUBJECT_START, min(IDX_SUBJECT_START + 10, df.shape[1])):
        subjects.append({
            "subject": subject_name_for_section(df, row_idx, j),
            "mark": format_percent(
                df.iloc[row_idx, j] if not pd.isna(df.iloc[row_idx, j]) else None
            ),
        })
    while len(subjects) < 10:
        subjects.append({"subject": "NotUploaded", "mark": "—"})

    return {
        "roll":       roll_display,
        "name":       name,
        "attendance": attendance,
        "subjects":   subjects,
    }


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="FCC Result Portal", page_icon="🎓", layout="wide")
st.title("🎓 Forman Christian College, Lahore")
st.title("Assessment & Examination Office")
st.subtitle("Facilitation Portal for Faculty, Staff, Students and Parents")
st.caption("Authenticate first, then enter a roll number to view overall attendance and all tests' results.")

# Startup: verify spreadsheet reachable
_sh, _err = _open_spreadsheet()
if _sh is None:
    st.error(f"❌ Cannot open spreadsheet `{SPREADSHEET_NAME}`: {_err}")
    st.stop()

st.session_state.setdefault("auth", False)
st.session_state.setdefault("email", "")
st.session_state.setdefault("roll", None)
st.session_state.setdefault("results", [])

with st.sidebar:
    st.caption("Source: Google Sheet")
    st.write(f"`{SPREADSHEET_NAME}`")
    tabs = available_tabs()
    st.write(f"Tabs: **{len(tabs)}**")
    st.caption(", ".join(tabs[:10]) + (" …" if len(tabs) > 10 else ""))
    st.write(f"Emails: **{len(accepted_emails())}**")

# ============================================================
# STEP 1 — AUTH
# ============================================================
st.header("Step 1 — Authenticate")

if not st.session_state.auth:
    with st.form("auth_form"):
        email_in = st.text_input(
            "Faculty email", placeholder="akramnaul@fccollege.edu.pk"
        )
        submitted = st.form_submit_button("Authenticate")

    if submitted:
        target = normalize(email_in)
        if not target:
            st.error("❌ Please enter an email.")
        else:
            accepted = accepted_emails()
            if target in accepted:
                log_auth_success(target)
                st.session_state.auth = True
                st.session_state.email = target
                st.rerun()
            else:
                log_auth_failure(target)
                st.error(f"❌ `{email_in.strip()}` not found in `{AUTH_TAB}`.")
                with st.expander("Show closest matches", expanded=True):
                    from difflib import get_close_matches
                    ms = get_close_matches(target, list(accepted), n=5, cutoff=0.5)
                    if ms:
                        for m in ms: st.code(m)
                    else:
                        st.write("No close matches.")
else:
    st.success(f"✅ Authenticated as: `{st.session_state.email}`")
    if st.button("Log out"):
        st.session_state.auth = False
        st.session_state.email = ""
        st.session_state.roll = None
        st.session_state.results = []
        st.rerun()

# ============================================================
# STEP 2 — SEARCH
# ============================================================
if st.session_state.auth:
    st.header("Step 2 — Enter Student Roll Number")
    
    with st.form("roll_form"):
        roll_raw = st.text_input(
            "Roll number",
            key="roll_input",
#            value=str(st.session_state.roll) if st.session_state.roll else "",
            placeholder="e.g. 277703101",
        )
        search = st.form_submit_button("Search Results")

    if search:
        roll = normalize_roll(roll_raw)
        if roll is None or roll <= 0:
            st.error(f"❌ `{roll_raw.strip()}` is not a valid roll number.")
        else:
            st.session_state.roll = roll
            log_search(st.session_state.email, roll)

            present = set(available_tabs())
            to_check = [t for t in RESULT_TABS if t in present]

            if not to_check:
                st.error(
                    f"❌ None of the configured result tabs exist. "
                    f"Expected: {RESULT_TABS}. Found: {sorted(present)}"
                )
                st.session_state.results = []
            else:
                found = []
                for tab in to_check:
                    df = load_tab_as_df(tab)
                    if df is None or df.empty:
                        continue
                    row = find_roll_row(df, roll)
                    if row is None:
                        continue
                    rec = extract_record(df, row)
                    rec["sheet"] = tab
                    rec["label"] = TAB_LABELS.get(tab, tab)
                    found.append(rec)

                st.session_state.results = found
                if not found:
                    st.error(f"❌ Roll `{roll}` not found in any result tab.")

    # ============================================================
    # STEP 3 — DISPLAY
    # ============================================================
    if st.session_state.results:
        st.header(
            f"📊 Results — `{st.session_state.roll}` : "
            f"`{st.session_state.results[0]['name']}`"
        )
        st.markdown("---")

        for rec in st.session_state.results:
            with st.container(border=True):
                st.subheader(f"🗓️ {rec['label']}  ({rec['sheet']})")
                st.markdown(
                    f"{rec['roll']} : {rec['name']} : "
                    f"**Overall Attendance** : {rec['attendance']}"
                )
                st.markdown("---")
                st.markdown("**Subjects & Marks (%):**")

                rows_html = ""
                for s in rec["subjects"]:
                    subj = s["subject"]
                    if len(subj) > 30:
                        subj = subj[:27] + "..."
                    rows_html += (
                        f"<tr>"
                        f"<td style='width:30ch;padding:2px 8px 2px 0;'>{subj}</td>"
                        f"<td style='padding:2px 0;font-weight:600;'>{s['mark']}</td>"
                        f"</tr>"
                    )
                st.markdown(
                    f"<table style='border-collapse:collapse;'>{rows_html}</table>",
                    unsafe_allow_html=True,
                )

        st.markdown("---")
        st.caption(
            "ℹ️ `NotUploaded` = section header missing for that column. "
            "`—` = cell was blank."
        )
