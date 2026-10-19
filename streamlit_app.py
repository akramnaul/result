# diagnostic.py  — save as streamlit_app.py temporarily
import streamlit as st
import pandas as pd
import re
from pathlib import Path

st.set_page_config(layout="wide", page_title="Roll Diagnostic")

SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_FILE = SCRIPT_DIR / "result.xlsx"

st.title("🔬 Roll Number Diagnostic")
st.write(f"Looking for: `{EXCEL_FILE}`  —  exists: **{EXCEL_FILE.exists()}**")

if not EXCEL_FILE.exists():
    st.error("File not found!")
    st.code(str(SCRIPT_DIR))
    st.code(sorted(p.name for p in SCRIPT_DIR.iterdir()))
    st.stop()

_INVIS = ("\u200b","\u200c","\u200d","\u2060","\ufeff","\xa0","\u2009","\u202f","\u00ad")

def parse_roll(v):
    if v is None: return None
    if isinstance(v, bool): return None
    if isinstance(v, float) and pd.isna(v): return None
    if isinstance(v, int): return v
    if isinstance(v, float): return int(v) if v.is_integer() else None
    s = str(v)
    for ch in _INVIS: s = s.replace(ch, "")
    s = s.strip().strip("'\"`‘’“”").strip()
    if not s: return None
    s = s.replace(",", "").replace(" ", "")
    try: return int(s)
    except ValueError: pass
    try:
        f = float(s); return int(f) if f.is_integer() else None
    except ValueError: pass
    d = re.sub(r"\D", "", s)
    return int(d) if d else None

target = st.text_input("Roll number to search", value="277703101")

with st.spinner("Reading workbook..."):
    xls = pd.ExcelFile(str(EXCEL_FILE))
    st.subheader(f"All sheet names ({len(xls.sheet_names)})")
    st.code("\n".join(xls.sheet_names))

    target_int = None
    try:
        target_int = int(target.strip())
    except ValueError:
        st.error(f"Cannot parse '{target}' as int")
        st.stop()

    st.write(f"Searching for integer: **{target_int}**")

    for sheet in xls.sheet_names:
        if sheet == "FacultyStaffList":
            continue  # skip

        st.markdown(f"### Sheet: `{sheet}`")

        # Read TWO ways to see what pandas sees
        df_header = pd.read_excel(str(EXCEL_FILE), sheet_name=sheet,
                                  header=0, dtype=object)
        df_nohdr  = pd.read_excel(str(EXCEL_FILE), sheet_name=sheet,
                                  header=None, dtype=object)

        st.write(f"With header=0 → shape `{df_header.shape}`")
        st.write(f"With header=None → shape `{df_nohdr.shape}`")

        # --- Column A as pandas with header=0 reads it ---
        col_a = df_header.iloc[:, 0] if df_header.shape[1] > 0 else pd.Series(dtype=object)
        non_null = col_a.dropna().tolist()

        st.write(f"**Column A non-null values (first 8, with repr):**")
        preview_lines = []
        matches_here = []
        for idx, v in enumerate(non_null):
            parsed = parse_roll(v)
            preview_lines.append(f"{v!r:<25}  ->  {parsed!r}")
            if parsed == target_int:
                # Find original DataFrame index
                orig_idx = col_a[col_a == v].index[0]
                matches_here.append((orig_idx, v))
            if idx >= 7:
                break
        st.code("\n".join(preview_lines) if preview_lines else "(empty)")

        # Full scan for the target
        all_parsed = col_a.map(parse_roll)
        hits = all_parsed[all_parsed == target_int]
        st.write(f"**Exact matches in column A: {len(hits)}**")
        if len(hits):
            for row_idx in hits.index[:3]:
                st.write(f"  → Row {row_idx}: raw value = `{col_a.iloc[row_idx]!r}`")

        # Try other columns too
        for j, cname in enumerate(df_header.columns[:6]):
            if j == 0: continue
            parsed_col = df_header.iloc[:, j].map(parse_roll)
            hits_j = parsed_col[parsed_col == target_int]
            if len(hits_j):
                st.write(f"  → **Also in column {chr(65+j)} (`{cname}`): "
                         f"{len(hits_j)} match(es)**")

        # Show ALL parsed values in column A so you can spot the target
        parsed_list = [p for p in all_parsed.tolist() if p is not None]
        if parsed_list:
            st.write(f"Sample parsed column A values: `{parsed_list[:10]}`")
        st.markdown("---")
