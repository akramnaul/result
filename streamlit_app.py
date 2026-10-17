# streamlit_app.py
"""
Student Monthly Result Portal.

Reads Result.xlsx from the SAME directory as this script (works locally and
on Streamlit Community Cloud where the repo is mounted at /mount/src/<repo>).

Sheet layout:
    - FacultyStaffList : column G holds faculty emails (may be prefixed with ')
    - Result-YY-MM     : A=Roll, B=Name, C=Attendance%, D..M=10 subjects (%)

All numeric/percentage values in columns C..M are displayed as whole-number
percentages (e.g. "87%" or "—" when empty). No fractions.
"""

import os
import re
from difflib import get_close_matches
from pathlib import Path

import pandas as pd
import streamlit as st

# ============================================================
# CONFIGURATION
# ============================================================
SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_FILE = SCRIPT_DIR / "result.xlsx"    # ⚠️ exact filename, case-sensitive

AUTH_SHEET        = "FacultyStaffList"
AUTH_EMAIL_COLUMN = "G"                    # Excel column letter (A=1, G=7)

# Ordered chronologically. Missing sheets are skipped automatically.
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

# ============================================================
# UTILITY
# ============================================================
def col_letter_to_index(letter: str) -> int:
    """'A' -> 0, 'B' -> 1, ..., 'AA' -> 26."""
    letter = letter.upper()
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def normalize(s) -> str:
    """
    Universal cell normalizer for emails and roll numbers.
    Strips leading/trailing apostrophes (Excel's force-text prefix), quotes,
    zero-width chars, BOM, NBSP, and surrounding whitespace. Lower-cases.
    """
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


_PCT_RE = re.compile(r"-?\d+(?:\.\d+)?")


def format_percent(raw) -> str:
    """
    Convert any cell value into a whole-number percentage string.

    Accepts:  None / "" / NaN      -> "—"
              "87"      (string)   -> "87%"
              "87%"     (string)   -> "87%"
              " 87.5% " (string)   -> "88%"   (rounds to nearest int)
              "0.87"    (fraction) -> "87%"   (converts to %)
              0.87      (float)    -> "87%"
              87        (int)      -> "87%"
              "87.5"    (numeric)  -> "88%"
    """
    if raw is None:
        return "—"
    if isinstance(raw, float) and pd.isna(raw):
        return "—"

    # Already a number (int/float) — Excel stored it as a number, not text.
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        val = float(raw)
        # If between 0 and 1 exclusive, assume it's a fraction → *100
        if 0.0 < val < 1.0:
            val *= 100.0
        return f"{round(val)}%"

    s = str(raw).replace("\xa0", " ").replace("\u200b", "").strip()
    if s == "":
        return "—"

    # Has an explicit percent sign → treat the number as already a percentage.
    if "%" in s:
        m = _PCT_RE.search(s.replace(",", ""))
        if not m:
            return "—"
        try:
            return f"{round(float(m.group(0)))}%"
        except ValueError:
            return "—"

    # No percent sign: try to parse as number.
    try:
        val = float(s.replace(",", ""))
    except ValueError:
        # Not numeric — return the raw string unchanged (e.g. "AB", "N/A")
        return s

    # If between 0 and 1 exclusive, treat as fraction → percent.
    if 0.0 < val < 1.0:
        val *= 100.0
    return f"{round(val)}%"


def _mtime() -> float:
    try:
        return EXCEL_FILE.stat().st_mtime
    except OSError:
        return 0.0


# ============================================================
# EXCEL READERS  (cached; auto-invalidated when file mtime changes)
# ============================================================
@st.cache_data(show_spinner=False)
def sheet_names(_m: float) -> list[str]:
    try:
        return pd.ExcelFile(str(EXCEL_FILE)).sheet_names
    except Exception:
        return []


@st.cache_data(show_spinner=False)
def read_sheet(sheet_name: str, _m: float):
    """Load a sheet with header row 0. Returns DataFrame or None."""
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
    """Load with header=None, for diagnostic display only."""
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
    """Set of accepted faculty emails, normalized (apostrophe/quote-safe)."""
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

    # Safety net: any column whose values mostly look like emails.
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


def subject_header(df: pd.DataFrame, col_letter: str) -> str:
    """
    Subject name is at the top of each ~50-student section.
    - Blank column (all empty)         -> 'NotUploaded'
    - First non-empty cell is numeric  -> 'NotUploaded' (header missing)
    - Otherwise                        -> that first non-empty cell
    """
    i = col_letter_to_index(col_letter)
    if i >= len(df.columns):
        return "NotUploaded"
    vals = df.iloc[:, i].dropna().astype(str).str.strip()
    vals = vals[vals != ""]
    if len(vals) == 0:
        return "NotUploaded"
    top = vals.iloc[0]
    try:
        float(top.replace(",", "").replace("%", "").strip())
        return "NotUploaded"
    except ValueError:
        return top


def find_roll_row(df: pd.DataFrame, roll: str):
    i = col_letter_to_index(ROLL_COL)
    if i >= len(df.columns):
        return None
    target = normalize(roll)
    rolls = df.iloc[:, i].map(normalize)
    hits = rolls[rolls == target]
    return hits.index[0] if len(hits) else None


def extract_record(df: pd.DataFrame, row_idx: int) -> dict:
    return {
        "roll":       safe_str(df, row_idx, ROLL_COL),
        "name":       safe_str(df, row_idx, NAME_COL),
        # attendance always rendered as whole-number percent
        "attendance": format_percent(raw_cell(df, row_idx, ATTENDANCE_COL)),
        "subjects": [
            {
                "subject": subject_header(df, c),
                # subject marks always rendered as whole-number percent
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
    st.markdown(
        "**Fix:** commit `Result.xlsx` to the same folder as "
        "`streamlit_app.py` in your GitHub repo (case-sensitive), "
        "then reboot the app on Streamlit Cloud."
    )
    st.stop()

CURRENT_MTIME = _mtime()

st.session_state.setdefault("auth", False)
st.session_state.setdefault("email", "")
st.session_state.setdefault("roll", "")
st.session_state.setdefault("results", [])

with st.sidebar:
    st.caption("Workbook")
    st.write(f"`{EXCEL_FILE.name}`")
    try:
        st.write(f"{EXCEL_FILE.stat().st_size/1024:.1f} KB")
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
            st.write(f"Column {AUTH_EMAIL_COLUMN} — {len(g_vals)} values "
                     f"(repr shows hidden characters):")
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
                st.session_state.auth = True
                st.session_state.email = target
                st.rerun()
            else:
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
                        st.write(
                            "No close matches. Check spelling and make sure "
                            "your email is entered in the workbook."
                        )
else:
    st.success(f"✅ Authenticated as: `{st.session_state.email}`")
    if st.button("Log out"):
        st.session_state.auth = False
        st.session_state.email = ""
        st.session_state.roll = ""
        st.session_state.results = []
        st.rerun()

# ============================================================
# STEP 2 — ROLL NUMBER LOOKUP
# ============================================================
if st.session_state.auth:
    st.header("Step 2 — Enter Student Roll Number")

    with st.form("roll_form"):
        roll_in = st.text_input("Roll number", value=st.session_state.roll)
        search = st.form_submit_button("Search Results")

    if search:
        roll = roll_in.strip()
        if not roll:
            st.error("❌ Please enter a roll number.")
        else:
            st.session_state.roll = roll
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
                left, right = st.columns(2)

                with left:
                    st.markdown("**📋 Student Info & Subject Names**")
                    st.write(f"**Roll No:** {rec['roll']}")
                    st.write(f"**Name:** {rec['name']}")
                    st.write(f"**Attendance:** {rec['attendance']}")
                    st.markdown("---")
                    st.markdown("**Subjects:**")
                    for s in rec["subjects"]:
                        st.write(f"- {s['subject']}")

                with right:
                    st.markdown("**📝 Marks (percentage)**")
                    st.write(f"**Name:** {rec['name']}")
                    st.write(f"**Attendance:** {rec['attendance']}")
                    st.markdown("---")
                    st.markdown("**Subject → Percentage:**")
                    for s in rec["subjects"]:
                        st.write(f"- **{s['subject']}:** {s['mark']}")

        st.markdown("---")
        st.caption(
            "ℹ️ All marks shown as whole-number percentages. "
            "`NotUploaded` means the subject teacher has not uploaded marks "
            "for that column yet. `—` means the cell was blank."
        )
