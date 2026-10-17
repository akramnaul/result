# streamlit_app.py
import os
from pathlib import Path
import pandas as pd
import streamlit as st

# ============================================================
# CONFIG
# ============================================================
SCRIPT_DIR   = Path(__file__).resolve().parent
EXCEL_FILE   = SCRIPT_DIR / "result.xlsx"   # ⚠️ exact filename on disk

AUTH_SHEET         = "FacultyStaffList"
AUTH_EMAIL_COLUMN  = "G"   # Excel column letter, 1-based: A=1, G=7

RESULT_SHEETS = [
    "Result-26-09", "Result-26-10", "Result-26-11", "Result-26-12",
    "Result-27-01", "Result-27-02", "Result-27-03", "Result-27-04",
    "Result-27-05", "Result-27-06", "Result-27-07", "Result-27-08",
]
SHEET_LABELS = {
    "Result-26-09": "September 2026", "Result-26-10": "October 2026",
    "Result-26-11": "November 2026", "Result-26-12": "December 2026",
    "Result-27-01": "January 2027",  "Result-27-02": "February 2027",
    "Result-27-03": "March 2027",    "Result-27-04": "April 2027",
    "Result-27-05": "May 2027",      "Result-27-06": "June 2027",
    "Result-27-07": "July 2027",     "Result-27-08": "August 2027",
}
ROLL_COL       = "A"
NAME_COL       = "B"
ATTENDANCE_COL = "C"
SUBJECT_COLS   = list("DEFGHIJKLM")

# ============================================================
# HELPERS
# ============================================================
def col_letter_to_index(letter: str) -> int:
    letter = letter.upper()
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _mtime() -> float:
    try:    return EXCEL_FILE.stat().st_mtime
    except OSError: return 0.0


@st.cache_data(show_spinner=False)
def read_sheet_raw(sheet_name: str, _m: float) -> pd.DataFrame | None:
    """Load a sheet with NO header assumption (header=None) for raw inspection."""
    try:
        return pd.read_excel(str(EXCEL_FILE), sheet_name=sheet_name,
                             header=None, dtype=str)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def read_sheet(sheet_name: str, _m: float) -> pd.DataFrame | None:
    """Load a sheet normally (first row as header)."""
    try:
        xls = pd.ExcelFile(str(EXCEL_FILE))
        if sheet_name not in xls.sheet_names:
            return None
        df = pd.read_excel(str(EXCEL_FILE), sheet_name=sheet_name,
                           header=0, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def sheet_names(_m: float) -> list[str]:
    try:    return pd.ExcelFile(str(EXCEL_FILE)).sheet_names
    except Exception: return []


def normalize_email(s: str) -> str:
    """Aggressive normalization: str, strip, lowercase, remove zero-width chars."""
    if s is None: return ""
    s = str(s)
    s = s.replace("\u200b", "").replace("\ufeff", "").replace("\xa0", " ")
    return s.strip().lower()


def collect_emails_from_column_g() -> set[str]:
    """
    Try several interpretations of 'Column G' and collect ALL email-looking
    values found. Returns set of normalized emails so auth succeeds no matter
    which interpretation matches the real file layout.
    """
    emails: set[str] = set()

    # ---- Interpretation 1: header=None, positional column G (index 6) ----
    raw = read_sheet_raw(AUTH_SHEET, _mtime())
    if raw is not None:
        g_idx = col_letter_to_index(AUTH_EMAIL_COLUMN)
        if g_idx < raw.shape[1]:
            for v in raw.iloc[:, g_idx].dropna():
                n = normalize_email(v)
                if "@" in n:
                    emails.add(n)

        # ---- Interpretation 2: header=None, ANY column that looks like email ----
        for c in range(raw.shape[1]):
            col_vals = raw.iloc[:, c].dropna()
            hits = sum(1 for v in col_vals if "@" in str(v))
            if hits >= 3 and hits >= len(col_vals) * 0.3:
                for v in col_vals:
                    n = normalize_email(v)
                    if "@" in n:
                        emails.add(n)

    # ---- Interpretation 3: header=0, column named like 'email' ----
    df = read_sheet(AUTH_SHEET, _mtime())
    if df is not None:
        for c in df.columns:
            if "email" in str(c).lower() or "mail" in str(c).lower():
                for v in df[c].dropna():
                    n = normalize_email(v)
                    if "@" in n:
                        emails.add(n)

    return emails


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="Student Result Portal", page_icon="🎓", layout="wide")
st.title("🎓 Student Monthly Result Portal")

# File presence
if not EXCEL_FILE.exists():
    st.error(f"❌ Excel file `{EXCEL_FILE.name}` not found at `{SCRIPT_DIR}`.")
    st.code(f"Files: {sorted(p.name for p in SCRIPT_DIR.iterdir())}")
    st.stop()

# Cache invalidation
_m = _mtime()

# ---------- DEBUG panel ----------
with st.expander("🔍 Debug — what the app sees in FacultyStaffList", expanded=True):
    st.write(f"File: `{EXCEL_FILE}`  •  size: {EXCEL_FILE.stat().st_size/1024:.1f} KB")
    st.write(f"Sheet names in workbook: `{sheet_names(_m)}`")

    raw = read_sheet_raw(AUTH_SHEET, _m)
    if raw is None:
        st.error(f"❌ Sheet '{AUTH_SHEET}' not found or unreadable.")
    else:
        st.write(f"`{AUTH_SHEET}` raw shape (header=None): **{raw.shape}**")
        st.write("First 8 rows, ALL columns (repr shows hidden whitespace):")
        st.dataframe(raw.head(8).applymap(lambda x: repr(x) if pd.notna(x) else x))

        g_idx = col_letter_to_index(AUTH_EMAIL_COLUMN)
        if g_idx < raw.shape[1]:
            g_vals = raw.iloc[:, g_idx].dropna().astype(str).tolist()
            st.write(f"Column **{AUTH_EMAIL_COLUMN}** (index {g_idx}) — "
                     f"{len(g_vals)} non-empty values, first 10 (repr'd):")
            st.code("\n".join(repr(v) for v in g_vals[:10]) or "(empty)")

        collected = collect_emails_from_column_g()
        st.write(f"Emails the app will accept: **{len(collected)}** unique")
        st.code("\n".join(sorted(collected)[:20]) or "(none detected)")

# ---------- Session ----------
st.session_state.setdefault("auth", False)
st.session_state.setdefault("email", "")
st.session_state.setdefault("roll", "")
st.session_state.setdefault("results", [])

# ---------- Step 1: Auth ----------
st.header("Step 1 — Authenticate")

if not st.session_state.auth:
    with st.form("auth"):
        email_in = st.text_input("Faculty email", placeholder="name@example.com")
        ok = st.form_submit_button("Authenticate")

    if ok:
        target = normalize_email(email_in)
        if not target:
            st.error("❌ Please enter an email.")
        else:
            accepted = collect_emails_from_column_g()
            if target in accepted:
                st.session_state.auth = True
                st.session_state.email = email_in.strip()
                st.rerun()
            else:
                st.error(f"❌ `{email_in.strip()}` is not in FacultyStaffList.")
                with st.expander("Why? Show closest matches", expanded=True):
                    from difflib import get_close_matches
                    matches = get_close_matches(target, list(accepted), n=5, cutoff=0.5)
                    if matches:
                        st.write("Did you mean:")
                        for m in matches:
                            st.code(m)
                    else:
                        st.write(
                            "No close matches. Confirm your email is spelled "
                            "exactly as in the workbook (see Debug panel above)."
                        )
else:
    st.success(f"✅ Authenticated as: {st.session_state.email}")
    if st.button("Log out"):
        for k in ("auth", "email", "roll", "results"):
            st.session_state[k] = "" if k != "auth" else False
            if k == "results": st.session_state[k] = []
        st.rerun()

# ---------- Step 2: Lookup ----------
if st.session_state.auth:
    st.header("Step 2 — Enter Roll Number")
    with st.form("roll"):
        roll_in = st.text_input("Roll number", value=st.session_state.roll)
        go = st.form_submit_button("Search Results")

    if go:
        roll = roll_in.strip()
        if not roll:
            st.error("❌ Enter a roll number.")
        else:
            st.session_state.roll = roll
            found = []
            for s in [x for x in RESULT_SHEETS if x in set(sheet_names(_m))]:
                df = read_sheet(s, _m)
                if df is None: continue
                g_idx = col_letter_to_index(ROLL_COL)
                if g_idx >= len(df.columns): continue
                rolls = df.iloc[:, g_idx].astype(str).str.strip()
                hit = rolls[rolls == roll]
                if len(hit) == 0: continue
                r = hit.index[0]

                def cell(letter):
                    i = col_letter_to_index(letter)
                    if i >= len(df.columns): return ""
                    v = df.iloc[r, i]
                    return "" if pd.isna(v) else str(v).strip()

                def header_for(letter):
                    i = col_letter_to_index(letter)
                    if i >= len(df.columns): return "NotUploaded"
                    vals = df.iloc[:, i].dropna().astype(str).str.strip()
                    vals = vals[vals != ""]
                    if len(vals) == 0: return "NotUploaded"
                    top = vals.iloc[0]
                    try: float(top.replace(",", "")); return "NotUploaded"
                    except ValueError: return top

                found.append({
                    "sheet": s, "label": SHEET_LABELS.get(s, s),
                    "roll": cell(ROLL_COL), "name": cell(NAME_COL),
                    "attendance": cell(ATTENDANCE_COL),
                    "subjects": [{"subject": header_for(c), "mark": cell(c)}
                                 for c in SUBJECT_COLS],
                })
            st.session_state.results = found
            if not found:
                st.error(f"❌ Roll `{roll}` not found in any result sheet.")

    # ---------- Step 3: Display ----------
    if st.session_state.results:
        st.header(f"📊 Results — Roll No: {st.session_state.roll}")
        st.subheader(f"Student: {st.session_state.results[0]['name']}")
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
                    for s in rec["subjects"]:
                        st.write(f"- {s['subject']}")
                with right:
                    st.markdown("**📝 Marks**")
                    st.write(f"**Name:** {rec['name']}")
                    st.write(f"**Attendance %:** {rec['attendance']}")
                    st.markdown("---")
                    for s in rec["subjects"]:
                        st.write(f"- **{s['subject']}:** {s['mark'] or '—'}")
