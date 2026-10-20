# streamlit_app.py
"""
FCC Assessment & Examination Portal — Student Monthly Result Viewer.

Reads result.xlsx (which must live beside this file) and displays all
monthly results for a given roll number after faculty authentication.
All authentication and search events are logged to the 'PortalActivity'
Google Sheet via activity_logger.py.
"""

import re
from pathlib import Path

import pandas as pd
import streamlit as st

from activity_logger import (
    log_auth_success,
    log_auth_failure,
    log_search,
)

# ============================================================
# CONFIG
# ============================================================
SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_FILE = SCRIPT_DIR / "result.xlsx"

AUTH_SHEET        = "FacultyStaffList"
AUTH_EMAIL_COLUMN = "G"

RESULT_SHEETS = [
    "Result-26-09", "Result-26-10", "Result-26-11", "Result-26-12",
    "Result-27-01", "Result-27-02", "Result-27-03", "Result-27-04",
    "Result-27-05", "Result-27-06", "Result-27-07", "Result-27-08",
]

SHEET_LABELS = {
    "Result-26-09": "September 2026", "Result-26-10": "October 2026",
    "Result-26-11": "November 2026",  "Result-26-12": "December 2026",
    "Result-27-01": "January 2027",   "Result-27-02": "February 2027",
    "Result-27-03": "March 2027",     "Result-27-04": "April 2027",
    "Result-27-05": "May 2027",       "Result-27-06": "June 2027",
    "Result-27-07": "July 2027",      "Result-27-08": "August 2027",
}

ROLL_COL       = "A"
NAME_COL       = "B"
ATTENDANCE_COL = "C"
SUBJECT_COLS   = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]

MAX_SECTION_LOOKBACK     = 50
MIN_HEADER_SUBJECT_CELLS = 3

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
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    s = str(value)
    for ch in _INVISIBLE:
        s = s.replace(ch, "")
    s = s.strip().strip("'\"`‘’“”").strip()
    return s.lower()


def normalize_roll(value):
    """Convert any roll representation into a Python int, or None."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else None

    s = str(value)
    for ch in _INVISIBLE:
        s = s.replace(ch, "")
    s = s.strip().strip("'\"`‘’“”").strip()
    if not s:
        return None

    s_clean = s.replace(",", "").replace(" ", "")
    try:
        return int(s_clean)
    except ValueError:
        pass
    try:
        f = float(s_clean)
        return int(f) if f.is_integer() else None
    except ValueError:
        pass
    digits = re.sub(r"\D", "", s)
    return int(digits) if digits else None


_PCT_RE = re.compile(r"-?\d+(?:\.\d+)?")


def format_percent(raw) -> str:
    if raw is None:
        return "—"
    if isinstance(raw, float) and pd.isna(raw):
        return "—"

    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        val = float(raw)
        if val == 1.0:
            return "100%"
        if 0.0 < val < 1.0:
            return f"{round(val * 100)}%"
        return f"{round(val)}%"

    s = str(raw)
    for ch in _INVISIBLE:
        s = s.replace(ch, "")
    s = s.strip()
    if not s:
        return "—"

    if "%" in s:
        m = _PCT_RE.search(s.replace(",", ""))
        if not m:
            return "—"
        try:
            return f"{round(float(m.group(0)))}%"
        except ValueError:
            return "—"

    try:
        val = float(s.replace(",", ""))
    except ValueError:
        return s

    if val == 1.0:
        return "100%"
    if 0.0 < val < 1.0:
        return f"{round(val * 100)}%"
    return f"{round(val)}%"


def _mtime() -> float:
    try:
        return EXCEL_FILE.stat().st_mtime
    except OSError:
        return 0.0


# ============================================================
# EXCEL READERS
# ============================================================
@st.cache_data(show_spinner=False)
def sheet_names(_m: float) -> list[str]:
    try:
        return pd.ExcelFile(str(EXCEL_FILE)).sheet_names
    except Exception:
        return []


@st.cache_data(show_spinner=False)
def read_sheet(sheet_name: str, _m: float):
    try:
        xls = pd.ExcelFile(str(EXCEL_FILE))
        if sheet_name not in xls.sheet_names:
            return None
        df = pd.read_excel(str(EXCEL_FILE), sheet_name=sheet_name,
                           header=0, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def read_sheet_raw(sheet_name: str, _m: float):
    try:
        return pd.read_excel(str(EXCEL_FILE), sheet_name=sheet_name,
                             header=None, dtype=object)
    except Exception:
        return None


# ============================================================
# AUTHENTICATION
# ============================================================
@st.cache_data(show_spinner=False)
def accepted_emails(_m: float) -> set[str]:
    emails: set[str] = set()
    raw = read_sheet_raw(AUTH_SHEET, _m)
    if raw is None:
        return emails

    g_idx = col_letter_to_index(AUTH_EMAIL_COLUMN)
    if g_idx < raw.shape[1]:
        for v in raw.iloc[:, g_idx].dropna():
            n = normalize(v)
            if "@" in n:
                emails.add(n)

    for c in range(raw.shape[1]):
        col_vals = raw.iloc[:, c].dropna().astype(str)
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
    """Search every column of df for the roll as an int."""
    target = int(roll)
    for j in range(df.shape[1]):
        parsed = df.iloc[:, j].map(normalize_roll)
        hits = parsed[parsed == target]
        if len(hits):
            return hits.index[0]
    return None


def find_roll_row_raw(sheet_name: str, roll: int, _m: float):
    """Fallback search on the raw (header=None) sheet."""
    target = int(roll)
    raw = read_sheet_raw(sheet_name, _m)
    if raw is None:
        return (None, None)

    for j in range(raw.shape[1]):
        parsed = raw.iloc[:, j].map(normalize_roll)
        hits = parsed[parsed == target]
        if len(hits):
            raw_row = hits.index[0]
            header_df_row = raw_row - 1
            if header_df_row >= 0:
                return (raw_row, header_df_row)
            return (raw_row, None)
    return (None, None)


# ============================================================
# SUBJECT NAME LOOKUP
# ============================================================
def find_section_header_row(df: pd.DataFrame, student_row: int,
                             roll_col_idx: int = 0):
    if roll_col_idx >= len(df.columns):
        return None
    subject_idx = [j for j in range(len(df.columns)) if j >= roll_col_idx + 3]
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


def subject_name_for_section(df: pd.DataFrame, student_row: int,
                              roll_col_idx: int, subject_col_idx: int) -> str:
    if subject_col_idx >= len(df.columns):
        return "NotUploaded"
    header_row = find_section_header_row(df, student_row, roll_col_idx)
    if header_row is None:
        return "NotUploaded"
    v = df.iloc[header_row, subject_col_idx]
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
    raw_roll = df.iloc[row_idx, 0] if df.shape[1] > 0 else None
    roll_int = normalize_roll(raw_roll)
    roll_display = (
        str(roll_int) if roll_int is not None
        else safe_str_cell(df, row_idx, 0)
    )

    subjects = []
    for j in range(3, min(13, df.shape[1])):
        subjects.append({
            "subject": subject_name_for_section(df, row_idx, 0, j),
            "mark": format_percent(
                df.iloc[row_idx, j] if not pd.isna(df.iloc[row_idx, j]) else None
            ),
        })
    while len(subjects) < 10:
        subjects.append({"subject": "NotUploaded", "mark": "—"})

    return {
        "roll": roll_display,
        "name": safe_str_cell(df, row_idx, 1),
        "attendance": format_percent(
            df.iloc[row_idx, 2]
            if df.shape[1] > 2 and not pd.isna(df.iloc[row_idx, 2])
            else None
        ),
        "subjects": subjects,
    }


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="FCC Result Portal", page_icon="🎓", layout="wide")
st.title("🎓 Forman Christian College, Lahore")
st.title("Assessment & Examination Office")
st.caption("Facilitation Portal for Faculty, Staff, Students and Parents")
st.caption(
    "Authenticate, then enter a roll number to view overall attendance "
    "and all tests' results."
)

if not EXCEL_FILE.exists():
    st.error(f"❌ Excel file `{EXCEL_FILE.name}` not found in `{SCRIPT_DIR}`")
    st.code(sorted(p.name for p in SCRIPT_DIR.iterdir()))
    st.stop()

CURRENT_MTIME = _mtime()

st.session_state.setdefault("auth", False)
st.session_state.setdefault("email", "")
st.session_state.setdefault("roll", None)
st.session_state.setdefault("results", [])
st.session_state.setdefault("search_diag", [])

with st.sidebar:
    st.caption("Workbook")
    st.write(f"`{EXCEL_FILE.name}`")
    try:
        st.write(f"{EXCEL_FILE.stat().st_size / 1024:.1f} KB")
    except OSError:
        pass
    st.write(f"Sheets: **{len(sheet_names(CURRENT_MTIME))}**")
    st.write(f"Emails: **{len(accepted_emails(CURRENT_MTIME))}**")

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
            accepted = accepted_emails(CURRENT_MTIME)
            if target in accepted:
                log_auth_success(target)
                st.session_state.auth = True
                st.session_state.email = target
                st.rerun()
            else:
                log_auth_failure(target)
                st.error(
                    f"❌ `{email_in.strip()}` not in FacultyStaffList column G."
                )
                with st.expander("Show closest matches", expanded=True):
                    from difflib import get_close_matches
                    ms = get_close_matches(
                        target, list(accepted), n=5, cutoff=0.5
                    )
                    if ms:
                        for m in ms:
                            st.code(m)
                    else:
                        st.write("No close matches.")
else:
    st.success(f"✅ Authenticated as: `{st.session_state.email}`")
    if st.button("Log out"):
        st.session_state.auth = False
        st.session_state.email = ""
        st.session_state.roll = None
        st.session_state.results = []
        st.session_state.search_diag = []
        st.rerun()

# ============================================================
# STEP 2 — SEARCH
# ============================================================
if st.session_state.auth:
    st.header("Step 2 — Enter Student Roll Number")

    with st.form("roll_form"):
        roll_raw = st.text_input(
            "Roll number",
            value=str(st.session_state.roll) if st.session_state.roll else "",
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

            existing = set(sheet_names(CURRENT_MTIME))
            to_check = [s for s in RESULT_SHEETS if s in existing]

            if not to_check:
                st.error("❌ No result sheets found.")
                st.session_state.results = []
                st.session_state.search_diag = []
            else:
                found = []
                diag = []
                for sheet in to_check:
                    df = read_sheet(sheet, CURRENT_MTIME)
                    if df is None:
                        diag.append(f"{sheet}: read failed")
                        continue

                    row = find_roll_row(df, roll)
                    if row is None:
                        raw_row, header_row = find_roll_row_raw(
                            sheet, roll, CURRENT_MTIME
                        )
                        if header_row is not None:
                            row = header_row
                            diag.append(
                                f"{sheet}: found via raw fallback at "
                                f"raw_row={raw_row}"
                            )
                        elif raw_row is not None:
                            diag.append(
                                f"{sheet}: roll found only in header row "
                                f"(raw_row=0)"
                            )
                            continue

                    if row is None:
                        col_summary = []
                        for j in range(df.shape[1]):
                            n = int(
                                (df.iloc[:, j].map(normalize_roll) == roll).sum()
                            )
                            if n:
                                col_summary.append(f"col{j}={n}")
                        diag.append(
                            f"{sheet}: no match ({df.shape}); "
                            f"{', '.join(col_summary) or 'target absent'}"
                        )
                        continue

                    rec = extract_record(df, row)
                    rec["sheet"] = sheet
                    rec["label"] = SHEET_LABELS.get(sheet, sheet)
                    found.append(rec)
                    diag.append(
                        f"{sheet}: FOUND at row {row} — "
                        f"roll={rec['roll']} name={rec['name']!r}"
                    )

                st.session_state.results = found
                st.session_state.search_diag = diag

                if not found:
                    st.error(
                        f"❌ Roll `{roll}` was not found in any result sheet."
                    )

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
                        f"<td style='width:30ch;padding:2px 8px 2px 0;'>"
                        f"{subj}</td>"
                        f"<td style='padding:2px 0;font-weight:600;'>"
                        f"{s['mark']}</td>"
                        f"</tr>"
                    )
                st.markdown(
                    f"<table style='border-collapse:collapse;'>"
                    f"{rows_html}</table>",
                    unsafe_allow_html=True,
                )

        st.markdown("---")
        st.caption(
            "ℹ️ `NotUploaded` = section header missing for that column. "
            "`—` = cell was blank."
        )
