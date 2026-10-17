# streamlit_app.py
import os
import base64
from pathlib import Path

import pandas as pd
import streamlit as st

# ============================================================
# CONFIGURATION
# ============================================================
# Resolve the directory that contains this script (works locally AND on
# Streamlit Community Cloud, where the repo is mounted at /mount/src/<repo>).
SCRIPT_DIR = Path(__file__).resolve().parent

# ⚠️ Filename on disk — must match EXACTLY (case-sensitive on Linux).

EXCEL_FILENAME = "result.xlsx"
EXCEL_FILE = SCRIPT_DIR / EXCEL_FILENAME

AUTH_SHEET = "FacultyStaffList"
AUTH_EMAIL_COLUMN = "G"          # Column G in FacultyStaffList holds emails

# Order matters: chronological. Missing sheets are skipped automatically.
RESULT_SHEETS = [
    "Result-26-09",
    "Result-26-10",
    "Result-26-11",
    "Result-26-12",
    "Result-27-01",
    "Result-27-02",
    "Result-27-03",
    "Result-27-04",
    "Result-27-05",
    "Result-27-06",
    "Result-27-07",
    "Result-27-08",
]

SHEET_LABELS = {
    "Result-26-09": "September 2026",
    "Result-26-10": "October 2026",
    "Result-26-11": "November 2026",
    "Result-26-12": "December 2026",
    "Result-27-01": "January 2027",
    "Result-27-02": "February 2027",
    "Result-27-03": "March 2027",
    "Result-27-04": "April 2027",
    "Result-27-05": "May 2027",
    "Result-27-06": "June 2027",
    "Result-27-07": "July 2027",
    "Result-27-08": "August 2027",
}

# Column layout (Excel letters)
ROLL_COL = "A"
NAME_COL = "B"
ATTENDANCE_COL = "C"
SUBJECT_COLS = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]

# ============================================================
# HELPERS
# ============================================================
def col_letter_to_index(letter: str) -> int:
    """Convert Excel column letter (A, B, ..., Z, AA, ...) to 0-based index."""
    letter = letter.upper()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _file_mtime() -> float:
    try:
        return EXCEL_FILE.stat().st_mtime
    except OSError:
        return 0.0


@st.cache_data(show_spinner=False)
def load_sheet(sheet_name: str, _mtime: float):
    """Load one sheet as a DataFrame, or None if it doesn't exist."""
    try:
        xls = pd.ExcelFile(str(EXCEL_FILE))
        if sheet_name not in xls.sheet_names:
            return None
        df = pd.read_excel(
            str(EXCEL_FILE), sheet_name=sheet_name, header=0, dtype=str
        )
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def list_sheet_names(_mtime: float):
    try:
        return pd.ExcelFile(str(EXCEL_FILE)).sheet_names
    except Exception:
        return []


def safe_cell(df: pd.DataFrame, row_idx: int, col_letter: str):
    col_idx = col_letter_to_index(col_letter)
    if col_idx >= len(df.columns):
        return None
    val = df.iloc[row_idx, col_idx]
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s else None


def authenticate_user(email: str) -> bool:
    df = load_sheet(AUTH_SHEET, _file_mtime())
    if df is None:
        return False
    col_idx = col_letter_to_index(AUTH_EMAIL_COLUMN)
    if col_idx >= len(df.columns):
        return False
    email_col = df.columns[col_idx]
    emails = (
        df[email_col].dropna().astype(str).str.strip().str.lower()
    )
    return email.strip().lower() in set(emails)


def find_roll_row(df: pd.DataFrame, roll: str):
    col_idx = col_letter_to_index(ROLL_COL)
    if col_idx >= len(df.columns):
        return None
    roll_col = df.columns[col_idx]
    rolls = df[roll_col].astype(str).str.strip()
    matches = rolls[rolls == str(roll).strip()]
    return matches.index[0] if len(matches) else None


def get_subject_header(df: pd.DataFrame, col_letter: str) -> str:
    """
    Subject name lives at the top of each ~50-student section.
    - If the entire column is blank → 'NotUploaded'
    - Otherwise → topmost non-empty cell is treated as the header.
      If it's a pure number, we still call it 'NotUploaded' (marks uploaded
      but header row is missing).
    """
    col_idx = col_letter_to_index(col_letter)
    if col_idx >= len(df.columns):
        return "NotUploaded"
    non_null = (
        df.iloc[:, col_idx].dropna().astype(str).str.strip()
    )
    non_null = non_null[non_null != ""]
    if len(non_null) == 0:
        return "NotUploaded"
    header = non_null.iloc[0]
    try:
        float(header.replace(",", ""))
        return "NotUploaded"
    except ValueError:
        return header


def extract_student_record(df: pd.DataFrame, row_idx: int):
    return {
        "roll": safe_cell(df, row_idx, ROLL_COL) or "",
        "name": safe_cell(df, row_idx, NAME_COL) or "",
        "attendance": safe_cell(df, row_idx, ATTENDANCE_COL) or "",
        "subjects": [
            {
                "subject": get_subject_header(df, c),
                "mark": safe_cell(df, row_idx, c) or "",
            }
            for c in SUBJECT_COLS
        ],
    }


def available_result_sheets():
    existing = set(list_sheet_names(_file_mtime()))
    return [s for s in RESULT_SHEETS if s in existing]


# ============================================================
# UI
# ============================================================
st.set_page_config(
    page_title="Student Result Portal", page_icon="🎓", layout="wide"
)
st.title("🎓 Student Monthly Result Portal")
st.caption(
    "Authenticate with your faculty email, then enter a roll number "
    "to view all available monthly results."
)

# ---------- Startup: verify the Excel file is present ----------
if not EXCEL_FILE.exists():
    st.error(f"❌ Excel file `{EXCEL_FILENAME}` not found.")
    st.markdown("**Diagnostics**")
    st.code(
        f"Script directory : {SCRIPT_DIR}\n"
        f"Expected file    : {EXCEL_FILE}\n"
        f"Current working  : {Path.cwd()}\n"
        f"Files in script dir: {sorted(p.name for p in SCRIPT_DIR.iterdir())}"
    )
    st.markdown(
        "**Fix:** commit the Excel file to the same folder as "
        "`streamlit_app.py` in your GitHub repository. Make sure the "
        "filename case matches exactly "
        f"(currently looking for `{EXCEL_FILENAME}`), and confirm the "
        "file is not excluded by `.gitignore` "
        "(`git add -f " + EXCEL_FILENAME + "`)."
    )
    st.stop()

# Optional bootstrap: if you ever want to expose the file for download
# (useful for testing that the correct bytes are on the server).
with st.sidebar:
    st.caption("About")
    st.write(f"Workbook: `{EXCEL_FILENAME}`")
    try:
        size_kb = EXCEL_FILE.stat().st_size / 1024
        st.write(f"Size: {size_kb:.1f} KB")
    except OSError:
        pass

# ---------- Session state ----------
for key, default in [
    ("authenticated", False),
    ("email", ""),
    ("roll", ""),
    ("results", []),
]:
    st.session_state.setdefault(key, default)

# ---------- STEP 1: Authentication ----------
st.header("Step 1 — Authenticate")

if not st.session_state.authenticated:
    with st.form("auth_form"):
        email_input = st.text_input(
            "Enter your faculty email address", placeholder="name@example.com"
        )
        submit_auth = st.form_submit_button("Authenticate")

    if submit_auth:
        if not email_input.strip():
            st.error("❌ Please enter an email address.")
        elif authenticate_user(email_input):
            st.session_state.authenticated = True
            st.session_state.email = email_input.strip()
            st.rerun()
        else:
            st.error(
                "❌ Email not found in `FacultyStaffList`. "
                "Please try again with the correct faculty email."
            )
else:
    st.success(f"✅ Authenticated as: {st.session_state.email}")
    if st.button("Log out"):
        st.session_state.authenticated = False
        st.session_state.email = ""
        st.session_state.roll = ""
        st.session_state.results = []
        st.rerun()

# ---------- STEP 2: Roll Number lookup ----------
if st.session_state.authenticated:
    st.header("Step 2 — Enter Student Roll Number")

    with st.form("roll_form"):
        roll_input = st.text_input(
            "Enter the student roll number", value=st.session_state.roll
        )
        submit_roll = st.form_submit_button("Search Results")

    if submit_roll:
        if not roll_input.strip():
            st.error("❌ Please enter a roll number.")
        else:
            st.session_state.roll = roll_input.strip()
            sheets = available_result_sheets()

            if not sheets:
                st.error(
                    "❌ No result sheets found in the workbook. "
                    "Expected sheets like `Result-26-09`, `Result-26-10`, …"
                )
                st.session_state.results = []
            else:
                found = []
                for sheet in sheets:
                    df = load_sheet(sheet, _file_mtime())
                    if df is None:
                        continue
                    row_idx = find_roll_row(df, st.session_state.roll)
                    if row_idx is None:
                        continue
                    rec = extract_student_record(df, row_idx)
                    rec["sheet"] = sheet
                    rec["label"] = SHEET_LABELS.get(sheet, sheet)
                    found.append(rec)

                if not found:
                    st.error(
                        f"❌ Roll number `{st.session_state.roll}` was not "
                        "found in any result sheet. Please try again with "
                        "the correct roll number."
                    )
                    st.session_state.results = []
                else:
                    st.session_state.results = found

    # ---------- STEP 3: Display results ----------
    if st.session_state.results:
        st.header(f"📊 Results for Roll No: {st.session_state.roll}")
        student_name = st.session_state.results[0]["name"]
        st.subheader(f"Student: {student_name}")
        st.markdown("---")

        for rec in st.session_state.results:
            with st.container(border=True):
                st.subheader(f"🗓️ {rec['label']}  ({rec['sheet']})")
                left, right = st.columns(2)

                with left:
                    st.markdown("**📋 Student Info & Subject Names**")
                    st.write(f"**Roll No:** {rec['roll']}")
                    st.write(f"**Name:** {rec['name']}")
                    st.write(f"**Attendance %:** {rec['attendance']}")
                    st.markdown("---")
                    st.markdown("**Subjects:**")
                    for s in rec["subjects"]:
                        st.write(f"- {s['subject']}")

                with right:
                    st.markdown("**📝 Marks**")
                    st.write(f"**Name:** {rec['name']}")
                    st.write(f"**Attendance %:** {rec['attendance']}")
                    st.markdown("---")
                    st.markdown("**Subject → Marks:**")
                    for s in rec["subjects"]:
                        mark = s["mark"] if s["mark"] != "" else "—"
                        subj = s["subject"]
                        st.write(f"- **{subj}:** {mark}")

        st.markdown("---")
        st.caption(
            "ℹ️ 'NotUploaded' means the subject teacher has not uploaded "
            "marks for that column yet (the column was blank)."
        )
