import st
import pandas as pd
import os
from pathlib import Path

# ============================================================
# CONFIGURATION
# ============================================================
SCRIPT_DIR = Path(__file__).resolve().parent

# ⚠️ Make sure this matches the actual filename on disk exactly
# (the spec says "result.xlsx", but the error text said "result.xlsx")
EXCEL_FILE = SCRIPT_DIR / "result.xlsx"

AUTH_SHEET = "FacultyStaffList"
AUTH_EMAIL_COLUMN = "G"

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

# Friendly month labels for display
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

# Column layout constants
ROLL_COL = "A"
NAME_COL = "B"
ATTENDANCE_COL = "C"
SUBJECT_COLS = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
NUM_SUBJECTS = len(SUBJECT_COLS)

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def col_letter_to_index(letter: str) -> int:
    """Convert Excel column letter (A, B, ... Z, AA) to 0-based index."""
    letter = letter.upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


@st.cache_data(show_spinner=False)
def load_excel_sheet(sheet_name: str, _file_mtime: float):
    """
    Load a sheet from the Excel file. `_file_mtime` is used only to invalidate cache
    when the file changes (Streamlit hashing trick).
    Returns DataFrame or None if sheet doesn't exist.
    """
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        if sheet_name not in xls.sheet_names:
            return None
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=0, dtype=str)
        # Normalize column names: strip whitespace
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except FileNotFoundError:
        return None
    except Exception:
        return None


def get_file_mtime():
    """Return file modification time for cache invalidation."""
    try:
        return os.path.getmtime(EXCEL_FILE)
    except OSError:
        return 0.0


def authenticate_user(email: str) -> bool:
    """
    Verify that the given email exists in FacultyStaffList column G.
    """
    mtime = get_file_mtime()
    df = load_excel_sheet(AUTH_SHEET, mtime)
    if df is None:
        return False

    col_idx = col_letter_to_index(AUTH_EMAIL_COLUMN)
    if col_idx >= len(df.columns):
        return False

    email_col = df.columns[col_idx]
    emails = df[email_col].dropna().astype(str).str.strip().str.lower()
    return email.strip().lower() in set(emails)


def safe_cell(df: pd.DataFrame, row_idx: int, col_letter: str):
    """Safely fetch a cell value from a DataFrame by Excel column letter."""
    col_idx = col_letter_to_index(col_letter)
    if col_idx >= len(df.columns):
        return None
    val = df.iloc[row_idx, col_idx]
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s != "" else None


def find_roll_row(df: pd.DataFrame, roll_number: str):
    """
    Search for roll number in column A.
    Returns row index (0-based integer position) or None.
    """
    col_idx = col_letter_to_index(ROLL_COL)
    if col_idx >= len(df.columns):
        return None

    roll_col = df.columns[col_idx]
    rolls = df[roll_col].astype(str).str.strip()
    target = str(roll_number).strip()
    matches = rolls[rolls == target]
    if len(matches) == 0:
        return None
    return matches.index[0]


def get_subject_header(df: pd.DataFrame, col_letter: str) -> str:
    """
    Determine the subject name for a given subject column.
    The hint says: "the name of each subject is available at the top of each
    section of about 50 students, or if blank, it means the numbers were not
    entered / uploaded for the subject's teacher. It should mark NotUploaded
    for that header."

    Interpretation:
    - Scan the column top-down for the first non-empty value that looks like a
      header/subject name (not a numeric mark).
    - If no non-empty values exist in that column at all (blank column), return
      "NotUploaded".
    - If values exist but they look like numbers only, still try the topmost
      non-empty cell as the header (per the "top of each section" hint).
    - If the topmost non-empty cell is numeric, we still treat the column as
      having content, but we can fall back to a generic label.
    """
    col_idx = col_letter_to_index(col_letter)
    if col_idx >= len(df.columns):
        return "NotUploaded"

    col = df.iloc[:, col_idx]
    non_null = col.dropna().astype(str).str.strip()
    non_null = non_null[non_null != ""]

    if len(non_null) == 0:
        return "NotUploaded"

    # Take the topmost non-empty value as the header (per hint)
    header = non_null.iloc[0].strip()

    # If it looks like a pure number, it's likely a mark — mark as unknown
    try:
        float(header.replace(",", ""))
        return "NotUploaded"
    except ValueError:
        return header


def extract_student_result(df: pd.DataFrame, row_idx: int):
    """
    Extract a student's record from a result sheet.
    Returns dict with roll, name, attendance, subjects list of (subject_name, mark).
    """
    roll = safe_cell(df, row_idx, ROLL_COL) or ""
    name = safe_cell(df, row_idx, NAME_COL) or ""
    attendance = safe_cell(df, row_idx, ATTENDANCE_COL) or ""

    subjects = []
    for col_letter in SUBJECT_COLS:
        subj_name = get_subject_header(df, col_letter)
        mark = safe_cell(df, row_idx, col_letter)
        subjects.append({
            "subject": subj_name,
            "mark": mark if mark is not None else "",
        })

    return {
        "roll": roll,
        "name": name,
        "attendance": attendance,
        "subjects": subjects,
    }


def get_available_result_sheets():
    """Return ordered list of result sheets that actually exist in the workbook."""
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        existing = set(xls.sheet_names)
    except Exception:
        return []
    return [s for s in RESULT_SHEETS if s in existing]


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="Student Result Portal", page_icon="🎓", layout="wide")

st.title("🎓 Student Monthly Result Portal")
st.caption("Authenticate with your faculty email, then enter a roll number to view all monthly results.")

# Session state
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "email" not in st.session_state:
    st.session_state.email = ""
if "roll" not in st.session_state:
    st.session_state.roll = ""
if "results" not in st.session_state:
    st.session_state.results = []

# ---------- File check ----------
if not Path(EXCEL_FILE).exists():
    st.error(f"❌ Excel file '{EXCEL_FILE}' not found in the application folder.")
    st.stop()

# ---------- STEP 1: Authentication ----------
st.header("Step 1 — Authenticate")

if not st.session_state.authenticated:
    with st.form("auth_form"):
        email_input = st.text_input("Enter your faculty email address", placeholder="name@example.com")
        auth_btn = st.form_submit_button("Authenticate")

    if auth_btn:
        if not email_input.strip():
            st.error("❌ Please enter an email address.")
        elif authenticate_user(email_input):
            st.session_state.authenticated = True
            st.session_state.email = email_input.strip()
            st.success(f"✅ Welcome, {email_input.strip()}!")
            st.rerun()
        else:
            st.error("❌ Email not found. Please try again with the correct faculty email address.")
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
        roll_input = st.text_input("Enter the student roll number", value=st.session_state.roll)
        search_btn = st.form_submit_button("Search Results")

    if search_btn:
        if not roll_input.strip():
            st.error("❌ Please enter a roll number.")
        else:
            st.session_state.roll = roll_input.strip()
            mtime = get_file_mtime()
            sheets = get_available_result_sheets()

            if not sheets:
                st.error(f"❌ No result sheets found in '{EXCEL_FILE}'.")
                st.session_state.results = []
            else:
                found_any = False
                results = []
                for sheet in sheets:
                    df = load_excel_sheet(sheet, mtime)
                    if df is None:
                        continue  # Hint #1: ignore missing sheet
                    row_idx = find_roll_row(df, st.session_state.roll)
                    if row_idx is not None:
                        found_any = True
                        record = extract_student_result(df, row_idx)
                        record["sheet"] = sheet
                        record["label"] = SHEET_LABELS.get(sheet, sheet)
                        results.append(record)

                if not found_any:
                    st.error(
                        f"❌ Roll number '{st.session_state.roll}' was not found in any result sheet. "
                        "Please try again with the correct roll number."
                    )
                    st.session_state.results = []
                else:
                    st.session_state.results = results

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

                # LEFT PANE: Roll, Name, Attendance, subject names vertically
                with left:
                    st.markdown("**📋 Student Info & Subject Names**")
                    st.write(f"**Roll No:** {rec['roll']}")
                    st.write(f"**Name:** {rec['name']}")
                    st.write(f"**Attendance %:** {rec['attendance']}")
                    st.write("---")
                    st.markdown("**Subjects:**")
                    for s in rec["subjects"]:
                        st.write(f"- {s['subject']}")

                # RIGHT PANE: attendance %, name and marks of 10 subjects vertically
                with right:
                    st.markdown("**📝 Marks**")
                    st.write(f"**Name:** {rec['name']}")
                    st.write(f"**Attendance %:** {rec['attendance']}")
                    st.write("---")
                    st.markdown("**Subject → Marks:**")
                    for s in rec["subjects"]:
                        mark_display = s["mark"] if s["mark"] != "" else "—"
                        subj_display = s["subject"] if s["subject"] != "NotUploaded" else "NotUploaded"
                        st.write(f"- **{subj_display}:** {mark_display}")

        st.markdown("---")
        st.caption(
            "ℹ️ 'NotUploaded' means the subject teacher has not uploaded marks yet "
            "for that column (the column was blank)."
        )
