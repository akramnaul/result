# streamlit_app.py
"""
Student Monthly Result Portal.

Reads result.xlsx from the SAME directory as this script (works locally and
on Streamlit Community Cloud where the repo is mounted at /mount/src/<repo>).

Sheet layout:
    - FacultyStaffList : column G holds faculty emails (may be prefixed with ')
                         A log is maintained starting at row 301:
                           B = Serial, C = EMailID, D = DateTime, E = Activity
                           Row 300 holds the column headings.
    - Result-YY-MM     : A=Roll, B=Name, C=Attendance%, D..M=10 subjects

Roll numbers:
    Treated as NUMBERS throughout. When the user enters a roll number, it is
    converted to an integer if possible, so it matches numeric cells in the
    result sheets and is written as a numeric value into the log.

Subject-name lookup:
    Subject names live in SECTION HEADER rows (rows whose Roll # cell in
    column A is empty). For each student, we scan upward (max 50 rows) to
    find their section header row and read the 10 subject names from there.
    If the header cell is empty, that subject's marks are NOT uploaded.

Logging:
    Every authentication attempt (success or failure) and every roll-number
    search is appended as a new row to FacultyStaffList, starting at row 301.

Percentage display:
    Columns C..M are shown as whole-number percentages (e.g. "87%") with "—"
    for empty cells. 100 is displayed as 100%, never 1%.
"""

import re
from datetime import datetime
from difflib import get_close_matches
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# ============================================================
# CONFIGURATION
# ============================================================
SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_FILE = SCRIPT_DIR / "result.xlsx"    # ⚠️ exact filename, case-sensitive

AUTH_SHEET        = "FacultyStaffList"
AUTH_EMAIL_COLUMN = "G"

# Logging configuration (1-based Excel rows/cols)
LOG_FIRST_DATA_ROW = 301
LOG_COL_SERIAL     = 2   # B
LOG_COL_EMAIL      = 3   # C
LOG_COL_DATETIME   = 4   # D
LOG_COL_ACTIVITY   = 5   # E

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

MAX_SECTION_LOOKBACK      = 50
MIN_HEADER_SUBJECT_CELLS  = 3

# ============================================================
# UTILITY
# ============================================================
def col_letter_to_index(letter: str) -> int:
    letter = letter.upper()
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def normalize(s) -> str:
    """Normalize emails (strip apostrophes, quotes, whitespace; lowercase)."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = (s.replace("\u200b", "")
          .replace("\ufeff", "")
          .replace("\xa0", " "))
    s = s.strip()
    s = s.strip("'\"`‘’“”")
    s = s.strip()
    return s.lower()


def normalize_roll(value):
    """
    Convert any roll-number representation into a Python int.

    Handles:  202830101 (int)       -> 202830101
              202830101.0 (float)   -> 202830101
              "202830101" (str)     -> 202830101
              "'202830101" (str)    -> 202830101  (Excel force-text apostrophe)
              " 202830101 "         -> 202830101
              202830101 (bool)      -> None       (bools are not roll numbers)

    Returns int, or None if the value cannot be interpreted as an integer.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, float) and pd.isna(value):
        return None

    # Already numeric
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return None

    # String: strip invisible chars, apostrophes, whitespace
    s = str(value)
    s = (s.replace("\u200b", "")
          .replace("\ufeff", "")
          .replace("\xa0", " "))
    s = s.strip().strip("'\"`‘’“”").strip()
    if s == "":
        return None

    try:
        # Try int first
        return int(s)
    except ValueError:
        pass

    # Try float (e.g. "202830101.0")
    try:
        f = float(s.replace(",", ""))
        if f.is_integer():
            return int(f)
        return None
    except ValueError:
        return None


_PCT_RE = re.compile(r"-?\d+(?:\.\d+)?")


def format_percent(raw) -> str:
    """Return whole-number percentage string, or '—' for empty."""
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

    s = str(raw).replace("\xa0", " ").replace("\u200b", "").strip()
    if s == "":
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
# LOGGING
# ============================================================
def log_activity(email: str, activity) -> None:
    """
    Append a row to FacultyStaffList starting at row 301.

    Column layout (1-based Excel columns):
        B = Serial (int)
        C = EmailID (text, prefixed with ' to force text)
        D = DateTime (text: "19 Oct 2026 21:14:51")
        E = Activity — either a string ("Authenticated",
                       "Authentication Failed") OR an INT roll number.
    """
    try:
        wb = load_workbook(str(EXCEL_FILE))
        if AUTH_SHEET not in wb.sheetnames:
            return
        ws = wb[AUTH_SHEET]

        last_row = LOG_FIRST_DATA_ROW - 1
        r = LOG_FIRST_DATA_ROW
        while r <= ws.max_row:
            v = ws.cell(row=r, column=LOG_COL_SERIAL).value
            if v is None or str(v).strip() == "":
                break
            last_row = r
            r += 1

        new_row = max(last_row + 1, LOG_FIRST_DATA_ROW)

        next_serial = 1
        if last_row >= LOG_FIRST_DATA_ROW:
            prev = ws.cell(row=last_row, column=LOG_COL_SERIAL).value
            try:
                next_serial = int(float(str(prev).strip())) + 1
            except (ValueError, TypeError):
                next_serial = last_row - LOG_FIRST_DATA_ROW + 2

        timestamp = datetime.now().strftime("%d %b %Y %H:%M:%S")

        # If activity is an int roll number, write as int. Otherwise as text.
        if isinstance(activity, int) and not isinstance(activity, bool):
            activity_value = activity
            activity_number_format = "0"      # integer format
        else:
            activity_value = str(activity)
            activity_number_format = "@"      # text format

        ws.cell(row=new_row, column=LOG_COL_SERIAL,   value=next_serial)
        ws.cell(row=new_row, column=LOG_COL_EMAIL,    value=f"'{email}")
        ws.cell(row=new_row, column=LOG_COL_DATETIME, value=timestamp)

        cell_activity = ws.cell(row=new_row, column=LOG_COL_ACTIVITY,
                                value=activity_value)
        cell_activity.number_format = activity_number_format
        cell_activity.alignment = Alignment(horizontal="left")

        ws.cell(row=new_row, column=LOG_COL_EMAIL).alignment = \
            Alignment(horizontal="left")
        ws.cell(row=new_row, column=LOG_COL_DATETIME).alignment = \
            Alignment(horizontal="left")

        wb.save(str(EXCEL_FILE))
    except Exception:
        pass


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
        df = pd.read_excel(
            str(EXCEL_FILE), sheet_name=sheet_name, header=0, dtype=object
        )
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def read_sheet_raw(sheet_name: str, _m: float):
    try:
        return pd.read_excel(
            str(EXCEL_FILE), sheet_name=sheet_name, header=None, dtype=object
        )
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
# SECTION-AWARE SUBJECT-NAME LOOKUP
# ============================================================
def find_section_header_row(df: pd.DataFrame, student_row: int):
    roll_idx = col_letter_to_index(ROLL_COL)
    if roll_idx >= len(df.columns):
        return None

    subject_idx = [
        col_letter_to_index(c) for c in SUBJECT_COLS
        if col_letter_to_index(c) < len(df.columns)
    ]

    start = student_row - 1
    stop  = max(0, student_row - MAX_SECTION_LOOKBACK)
    for r in range(start, stop - 1, -1):
        v = df.iloc[r, roll_idx]
        if pd.isna(v) or str(v).strip() == "":
            filled = sum(
                1 for i in subject_idx
                if not pd.isna(df.iloc[r, i]) and str(df.iloc[r, i]).strip() != ""
            )
            if filled >= MIN_HEADER_SUBJECT_CELLS:
                return r
    return None


def subject_name_for_section(df: pd.DataFrame, student_row: int,
                              col_letter: str) -> str:
    i = col_letter_to_index(col_letter)
    if i >= len(df.columns):
        return "NotUploaded"

    header_row = find_section_header_row(df, student_row)
    if header_row is None:
        return "NotUploaded"

    v = df.iloc[header_row, i]
    if pd.isna(v):
        return "NotUploaded"
    s = str(v).strip()
    if s == "":
        return "NotUploaded"
    try:
        float(s.replace(",", "").replace("%", "").strip())
        return "NotUploaded"
    except ValueError:
        return s


# ============================================================
# RESULT LOOKUP
# ============================================================
def raw_cell(df: pd.DataFrame, row_idx: int, col_letter: str):
    i = col_letter_to_index(col_letter)
    if i >= len(df.columns):
        return None
    v = df.iloc[row_idx, i]
    return None if pd.isna(v) else v


def safe_str(df: pd.DataFrame, row_idx: int, col_letter: str) -> str:
    v = raw_cell(df, row_idx, col_letter)
    return "" if v is None else str(v).strip()


def find_roll_row(df: pd.DataFrame, roll: int):
    """
    Find the row where column A (as an int) matches `roll` (also int).
    Rolls stored in the sheet as "'202830101" or "202830101.0" are correctly
    converted to int for the comparison.
    """
    i = col_letter_to_index(ROLL_COL)
    if i >= len(df.columns):
        return None

    rolls_int = df.iloc[:, i].map(normalize_roll)
    hits = rolls_int[rolls_int == roll]
    return hits.index[0] if len(hits) else None


def extract_record(df: pd.DataFrame, row_idx: int) -> dict:
    # Display the roll number as an integer string (no .0 suffix, no apostrophe)
    raw_roll = raw_cell(df, row_idx, ROLL_COL)
    roll_int = normalize_roll(raw_roll)
    roll_display = str(roll_int) if roll_int is not None else safe_str(df, row_idx, ROLL_COL)

    return {
        "roll":       roll_display,
        "name":       safe_str(df, row_idx, NAME_COL),
        "attendance": format_percent(raw_cell(df, row_idx, ATTENDANCE_COL)),
        "subjects": [
            {
                "subject": subject_name_for_section(df, row_idx, c),
                "mark":    format_percent(raw_cell(df, row_idx, c)),
            }
            for c in SUBJECT_COLS
        ],
    }


# ============================================================
# UI SETUP
# ============================================================
st.set_page_config(
    page_title="Student Result Portal",
    page_icon="🎓",
    layout="wide",
)
st.title("🎓 Student Monthly Result Portal")
st.caption(
    "Authenticate with your faculty email, then enter a student roll number "
    "to view all available monthly results."
)

if not EXCEL_FILE.exists():
    st.error(f"❌ Excel file `{EXCEL_FILE.name}` was not found.")
    st.markdown("**Diagnostics**")
    st.code(
        f"Script directory     : {SCRIPT_DIR}\n"
        f"Expected file        : {EXCEL_FILE}\n"
        f"Current working dir  : {Path.cwd()}\n"
        f"Files in script dir  : "
        f"{sorted(p.name for p in SCRIPT_DIR.iterdir())}"
    )
    st.stop()

CURRENT_MTIME = _mtime()

st.session_state.setdefault("auth", False)
st.session_state.setdefault("email", "")
st.session_state.setdefault("roll", None)
st.session_state.setdefault("results", [])

with st.sidebar:
    st.caption("Workbook")
    st.write(f"`{EXCEL_FILE.name}`")
    try:
        st.write(f"{EXCEL_FILE.stat().st_size / 1024:.1f} KB")
    except OSError:
        pass
    st.write(f"Emails on file: **{len(accepted_emails(CURRENT_MTIME))}**")

with st.expander("🔍 Debug — what the app sees", expanded=False):
    st.write(f"Sheet names: `{sheet_names(CURRENT_MTIME)}`")
    raw = read_sheet_raw(AUTH_SHEET, CURRENT_MTIME)
    if raw is None:
        st.error(f"Sheet `{AUTH_SHEET}` not found.")
    else:
        g_idx = col_letter_to_index(AUTH_EMAIL_COLUMN)
        if g_idx < raw.shape[1]:
            g_vals = raw.iloc[:, g_idx].dropna().astype(str).tolist()
            st.write(f"Column {AUTH_EMAIL_COLUMN} — {len(g_vals)} values:")
            st.code("\n".join(repr(v) for v in g_vals[:10]) or "(empty)")
        st.write(f"Accepted emails (normalized): "
                 f"**{len(accepted_emails(CURRENT_MTIME))}**")

# ============================================================
# STEP 1 — AUTHENTICATION
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
            st.error("❌ Please enter an email address.")
        else:
            accepted = accepted_emails(CURRENT_MTIME)
            if target in accepted:
                log_activity(target, "Authenticated")
                st.session_state.auth = True
                st.session_state.email = target
                st.rerun()
            else:
                log_activity(target, "Authentication Failed")
                st.error(
                    f"❌ `{email_in.strip()}` was not found in "
                    f"`{AUTH_SHEET}` column {AUTH_EMAIL_COLUMN}."
                )
                with st.expander("Show closest matches", expanded=True):
                    matches = get_close_matches(
                        target, list(accepted), n=5, cutoff=0.5
                    )
                    if matches:
                        st.write("Did you mean:")
                        for m in matches:
                            st.code(m)
                    else:
                        st.write("No close matches. Check spelling.")
else:
    st.success(f"✅ Authenticated as: `{st.session_state.email}`")
    if st.button("Log out"):
        st.session_state.auth = False
        st.session_state.email = ""
        st.session_state.roll = None
        st.session_state.results = []
        st.rerun()

# ============================================================
# STEP 2 — ROLL NUMBER LOOKUP
# ============================================================
if st.session_state.auth:
    st.header("Step 2 — Enter Student Roll Number")

    with st.form("roll_form"):
        # Show as a numeric field (no letters accepted)
        roll_in = st.number_input(
            "Roll number",
            min_value=0,
            step=1,
            value=int(st.session_state.roll) if st.session_state.roll else 0,
            format="%d",
        )
        search = st.form_submit_button("Search Results")

    if search:
        roll = int(roll_in)
        if roll <= 0:
            st.error("❌ Please enter a valid roll number.")
        else:
            st.session_state.roll = roll

            # Log the search as a numeric activity
            log_activity(st.session_state.email, roll)

            existing = set(sheet_names(CURRENT_MTIME))
            to_check = [s for s in RESULT_SHEETS if s in existing]

            if not to_check:
                st.error(
                    "❌ No result sheets found in the workbook. "
                    "Expected sheets like `Result-26-09`, `Result-26-10`, …"
                )
                st.session_state.results = []
            else:
                found = []
                for sheet in to_check:
                    df = read_sheet(sheet, CURRENT_MTIME)
                    if df is None:
                        continue
                    row = find_roll_row(df, roll)
                    if row is None:
                        continue
                    rec = extract_record(df, row)
                    rec["sheet"] = sheet
                    rec["label"] = SHEET_LABELS.get(sheet, sheet)
                    found.append(rec)

                if not found:
                    st.error(
                        f"❌ Roll number `{roll}` was not found in any "
                        "result sheet. Please try again with the correct "
                        "roll number."
                    )
                    st.session_state.results = []
                else:
                    st.session_state.results = found

    # ============================================================
    # STEP 3 — DISPLAY RESULTS
    # ============================================================
    if st.session_state.results:
        st.header(f"📊 Results — Roll No: `{st.session_state.roll}`")
        st.subheader(f"Student: {st.session_state.results[0]['name']}")
        st.markdown("---")

        for rec in st.session_state.results:
            with st.container(border=True):
                st.subheader(f"🗓️ {rec['label']}  ({rec['sheet']})")

                c_left, c_right = st.columns(2)
                with c_left:
                    st.markdown(f"**Roll No:** {rec['roll']}")
                    st.markdown(f"**Name:** {rec['name']}")

                st.markdown(f"**Attendance:** {rec['attendance']}")
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
            "ℹ️ All marks shown as whole-number percentages. "
            "`NotUploaded` means the subject teacher has not uploaded marks "
            "for that column yet. `—` means the cell was blank."
        )
