# streamlit_app.py
import re
from datetime import datetime
from difflib import get_close_matches
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# ============================================================
# CONFIG
# ============================================================
SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_FILE = SCRIPT_DIR / "result.xlsx"

AUTH_SHEET        = "FacultyStaffList"
AUTH_EMAIL_COLUMN = "G"

LOG_FIRST_DATA_ROW = 301
LOG_COL_SERIAL     = 2
LOG_COL_EMAIL      = 3
LOG_COL_DATETIME   = 4
LOG_COL_ACTIVITY   = 5

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

# After reading with header=0, roll is column A (index 0). But if a sheet
# has an extra title row, the data shifts. We'll detect columns dynamically.

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
    if value is None: return ""
    if isinstance(value, float) and pd.isna(value): return ""
    s = str(value)
    for ch in _INVISIBLE: s = s.replace(ch, "")
    s = s.strip().strip("'\"`‘’“”").strip()
    return s.lower()


def normalize_roll(value):
    """Convert ANY roll representation to int. Returns int or None."""
    if value is None: return None
    if isinstance(value, bool): return None
    if isinstance(value, float) and pd.isna(value): return None
    if isinstance(value, int): return value
    if isinstance(value, float): return int(value) if value.is_integer() else None

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

    # Digit-only fallback
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


def _mtime() -> float:
    try: return EXCEL_FILE.stat().st_mtime
    except OSError: return 0.0


# ============================================================
# LOGGING
# ============================================================
def log_activity(email: str, activity) -> None:
    try:
        wb = load_workbook(str(EXCEL_FILE))
        if AUTH_SHEET not in wb.sheetnames: return
        ws = wb[AUTH_SHEET]

        last_row = LOG_FIRST_DATA_ROW - 1
        r = LOG_FIRST_DATA_ROW
        while r <= ws.max_row:
            v = ws.cell(row=r, column=LOG_COL_SERIAL).value
            if v is None or str(v).strip() == "": break
            last_row = r
            r += 1

        new_row = max(last_row + 1, LOG_FIRST_DATA_ROW)
        next_serial = 1
        if last_row >= LOG_FIRST_DATA_ROW:
            prev = ws.cell(row=last_row, column=LOG_COL_SERIAL).value
            try: next_serial = int(float(str(prev).strip())) + 1
            except (ValueError, TypeError):
                next_serial = last_row - LOG_FIRST_DATA_ROW + 2

        timestamp = datetime.now().strftime("%d %b %Y %H:%M:%S")

        if isinstance(activity, int) and not isinstance(activity, bool):
            activity_value, activity_format = activity, "0"
        else:
            activity_value, activity_format = str(activity), "@"

        ws.cell(row=new_row, column=LOG_COL_SERIAL,   value=next_serial)
        ws.cell(row=new_row, column=LOG_COL_EMAIL,    value=f"'{email}")
        ws.cell(row=new_row, column=LOG_COL_DATETIME, value=timestamp)
        c_act = ws.cell(row=new_row, column=LOG_COL_ACTIVITY, value=activity_value)
        c_act.number_format = activity_format
        c_act.alignment = Alignment(horizontal="left")
        ws.cell(row=new_row, column=LOG_COL_EMAIL).alignment = Alignment(horizontal="left")
        ws.cell(row=new_row, column=LOG_COL_DATETIME).alignment = Alignment(horizontal="left")
        wb.save(str(EXCEL_FILE))
    except Exception:
        pass


# ============================================================
# EXCEL READERS
# ============================================================
@st.cache_data(show_spinner=False)
def sheet_names(_m: float) -> list[str]:
    try: return pd.ExcelFile(str(EXCEL_FILE)).sheet_names
    except Exception: return []


@st.cache_data(show_spinner=False)
def read_sheet(sheet_name: str, _m: float):
    try:
        xls = pd.ExcelFile(str(EXCEL_FILE))
        if sheet_name not in xls.sheet_names: return None
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
    if raw is None: return emails

    g_idx = col_letter_to_index(AUTH_EMAIL_COLUMN)
    if g_idx < raw.shape[1]:
        for v in raw.iloc[:, g_idx].dropna():
            n = normalize(v)
            if "@" in n: emails.add(n)

    for c in range(raw.shape[1]):
        col_vals = raw.iloc[:, c].dropna().astype(str)
        if len(col_vals) == 0: continue
        hits = sum(1 for v in col_vals if "@" in v)
        if hits >= 3 and hits >= len(col_vals) * 0.3:
            for v in col_vals:
                n = normalize(v)
                if "@" in n: emails.add(n)
    return emails


# ============================================================
# ROLL DISCOVERY — most robust version
# ============================================================
def find_roll_row(df: pd.DataFrame, roll: int):
    """
    Search for `roll` (as int) across ALL columns of df.
    Returns the row index (0-based) if found, else None.
    Because the diagnostic already proved the number is present somewhere,
    we scan every column to be immune to layout differences.
    """
    target = int(roll)
    for j in range(df.shape[1]):
        col = df.iloc[:, j]
        parsed = col.map(normalize_roll)
        hits = parsed[parsed == target]
        if len(hits):
            return hits.index[0]
    return None


def find_roll_row_raw(sheet_name: str, roll: int, _m: float):
    """
    Fallback: search the RAW sheet (header=None) for the roll. Used when
    find_roll_row(df) fails, to catch cases where a title row shifted data.
    Returns (row_idx, header0_df_row_idx) or (None, None).
    """
    target = int(roll)
    raw = read_sheet_raw(sheet_name, _m)
    if raw is None: return (None, None)

    for j in range(raw.shape[1]):
        col = raw.iloc[:, j]
        parsed = col.map(normalize_roll)
        hits = parsed[parsed == target]
        if len(hits):
            raw_row = hits.index[0]
            # Shift by -1 because raw has one extra row (no header consumed)
            header_df_row = raw_row - 1
            if header_df_row >= 0:
                return (raw_row, header_df_row)
            return (raw_row, None)
    return (None, None)


# ============================================================
# SUBJECT NAME
# ============================================================
def find_section_header_row(df: pd.DataFrame, student_row: int,
                             roll_col_idx: int = 0):
    if roll_col_idx >= len(df.columns): return None
    subject_idx = [j for j in range(len(df.columns)) if j >= roll_col_idx + 3]
    start = student_row - 1
    stop  = max(0, student_row - MAX_SECTION_LOOKBACK)
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
    if subject_col_idx >= len(df.columns): return "NotUploaded"
    header_row = find_section_header_row(df, student_row, roll_col_idx)
    if header_row is None: return "NotUploaded"
    v = df.iloc[header_row, subject_col_idx]
    if pd.isna(v): return "NotUploaded"
    s = str(v).strip()
    if not s: return "NotUploaded"
    try:
        float(s.replace(",", "").replace("%", "").strip())
        return "NotUploaded"
    except ValueError:
        return s


# ============================================================
# RECORD EXTRACTION
# ============================================================
def safe_str_cell(df, row_idx, col_idx):
    if col_idx >= len(df.columns): return ""
    v = df.iloc[row_idx, col_idx]
    return "" if pd.isna(v) else str(v).strip()


def extract_record(df: pd.DataFrame, row_idx: int) -> dict:
    """Assumes a standard layout: col 0=Roll, col 1=Name, col 2=Attendance,
    cols 3..12=subject marks. Subject names come from the nearest section
    header row above."""
    raw_roll = df.iloc[row_idx, 0] if df.shape[1] > 0 else None
    roll_int = normalize_roll(raw_roll)
    roll_display = str(roll_int) if roll_int is not None else safe_str_cell(df, row_idx, 0)

    subjects = []
    for j in range(3, min(13, df.shape[1])):
        subjects.append({
            "subject": subject_name_for_section(df, row_idx, 0, j),
            "mark":    format_percent(df.iloc[row_idx, j] if not pd.isna(df.iloc[row_idx, j]) else None),
        })
    # If fewer than 10 subject columns were present, pad with NotUploaded rows
    while len(subjects) < 10:
        subjects.append({"subject": "NotUploaded", "mark": "—"})

    return {
        "roll":       roll_display,
        "name":       safe_str_cell(df, row_idx, 1),
        "attendance": format_percent(df.iloc[row_idx, 2] if df.shape[1] > 2 and not pd.isna(df.iloc[row_idx, 2]) else None),
        "subjects":   subjects,
    }


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="FCC Result Portal", page_icon="🎓", layout="wide")
st.title("🎓 Forman Christian College, Lahore")
st.title("Assessment & Examination Office")
st.caption("Facilitation Portal for Faculty, Staff, Students and Parents")
st.caption("Authenticate, then enter a roll number to view his overall attendance and all tests' results.")

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
    try: st.write(f"{EXCEL_FILE.stat().st_size/1024:.1f} KB")
    except OSError: pass
    st.write(f"Sheets: **{len(sheet_names(CURRENT_MTIME))}**")
    st.write(f"Emails: **{len(accepted_emails(CURRENT_MTIME))}**")

# ============================================================
# STEP 1 — AUTH
# ============================================================
st.header("Step 1 — Authenticate")

if not st.session_state.auth:
    with st.form("auth_form"):
        email_in = st.text_input("Faculty email", placeholder="akramnaul@fccollege.edu.pk")
        submitted = st.form_submit_button("Authenticate")

    if submitted:
        target = normalize(email_in)
        if not target:
            st.error("❌ Please enter an email.")
        else:
            accepted = accepted_emails(CURRENT_MTIME)
            if target in accepted:
                log_activity(target, "Authenticated")
                st.session_state.auth = True
                st.session_state.email = target
                st.rerun()
            else:
                log_activity(target, "Authentication Failed")
                st.error(f"❌ `{email_in.strip()}` not in FacultyStaffList column G.")
                with st.expander("Show closest matches", expanded=True):
                    ms = get_close_matches(target, list(accepted), n=5, cutoff=0.5)
                    if ms:
                        for m in ms: st.code(m)
                    else:
                        st.write("No close matches.")
else:
    st.success(f"✅ Authenticated as: `{st.session_state.email}`")
    if st.button("Log out"):
        for k in ("auth","email","roll","results","search_diag"):
            st.session_state[k] = False if k=="auth" else ("" if k=="email" else (None if k=="roll" else []))
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
            log_activity(st.session_state.email, roll)

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

                    # Try header=0 DataFrame first
                    row = find_roll_row(df, roll)
                    if row is None:
                        # Fallback: search raw (header=None), shift row by -1
                        raw_row, header_row = find_roll_row_raw(sheet, roll, CURRENT_MTIME)
                        if header_row is not None:
                            row = header_row
                            diag.append(f"{sheet}: found via raw fallback at raw_row={raw_row}")
                        elif raw_row is not None:
                            # Roll is in raw row 0 → must be a header line, no data row
                            diag.append(f"{sheet}: roll found only in header row (raw_row=0)")
                            continue

                    if row is None:
                        # Column-scan diagnostic
                        col_summary = []
                        for j in range(df.shape[1]):
                            n = int((df.iloc[:, j].map(normalize_roll) == roll).sum())
                            if n: col_summary.append(f"col{j}={n}")
                        diag.append(f"{sheet}: no match ({df.shape}); {', '.join(col_summary) or 'target absent'}")
                        continue

                    rec = extract_record(df, row)
                    rec["sheet"] = sheet
                    rec["label"] = SHEET_LABELS.get(sheet, sheet)
                    found.append(rec)
                    diag.append(f"{sheet}: FOUND at row {row} — roll={rec['roll']} name={rec['name']!r}")

                st.session_state.results = found
                st.session_state.search_diag = diag

                if not found:
                    st.error(f"❌ Roll `{roll}` was not found in any result sheet.")

    # --- Diagnostic panel (always available) ---
#    if st.session_state.search_diag:
#        with st.expander("🔍 Search diagnostic", expanded=True):
#            st.code("\n".join(st.session_state.search_diag))

    # ============================================================
    # STEP 3 — DISPLAY
    # ============================================================
    if st.session_state.results:
        st.header(f"📊 Results — `{st.session_state.roll}` : `{st.session_state.results[0]['name']}`")
#        st.subheader(f"Student: {st.session_state.results[0]['name']}")
#        st.header(f"📊 Results — Roll No: `{st.session_state.roll}`")
#        st.subheader(f"Student: {st.session_state.results[0]['name']}")
        st.markdown("---")

        for rec in st.session_state.results:
            with st.container(border=True):
                st.subheader(f"🗓️ {rec['label']}  ({rec['sheet']})")
                c_left, c_right = st.columns(2)
                with c_left:
                    st.markdown(f"{rec['roll']} : {rec['name']} : **Overall Attendance** : {rec['attendance']}")
#                    st.markdown(f"**Roll No:** {rec['roll']}")
#                    st.markdown(f"**Name:** {rec['name']}")
#                st.markdown(f"**Attendance:** {rec['attendance']}")
                st.markdown("---")
                st.markdown("**Subjects & Marks (%):**")

                rows_html = ""
                for s in rec["subjects"]:
                    subj = s["subject"]
                    if len(subj) > 30: subj = subj[:27] + "..."
                    rows_html += (
                        f"<tr>"
                        f"<td style='width:30ch;padding:2px 8px 2px 0;'>{subj}</td>"
                        f"<td style='padding:2px 0;font-weight:600;'>{s['mark']}</td>"
                        f"</tr>"
                    )
                st.markdown(f"<table style='border-collapse:collapse;'>{rows_html}</table>",
                            unsafe_allow_html=True)

        st.markdown("---")
        st.caption("ℹ️ `NotUploaded` = section header missing for that column. "
                   "`—` = cell was blank.")
