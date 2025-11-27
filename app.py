import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Security Group Access Comparator", layout="wide")
st.title("🔎 Security Group Access Comparator")
st.markdown("Upload **Standard** and **Client** Excel files (same headers). App shows 3 findings and lets you download results.")

# --- Helpers ---
SG_COL = "Domains granted to Security Group"

def normalize_text_cell(cell):
    """Return list of stripped non-empty lines from a cell (handles NaN)."""
    if pd.isna(cell):
        return []
    if isinstance(cell, (int, float)):
        cell = str(cell)
    # split on newline, also handle semicolon/comma if needed in future
    parts = [p.strip() for p in str(cell).splitlines() if p.strip() != ""]
    return parts

def compare_cells(std_cell, clt_cell):
    """Return lists: missing (in std not in clt), extra (in clt not in std)"""
    sset = set(normalize_text_cell(std_cell))
    cset = set(normalize_text_cell(clt_cell))
    missing = sorted(list(sset - cset))
    extra = sorted(list(cset - sset))
    return missing, extra

def build_difference_text(missing, extra, use_html=False):
    """
    Build a multi-line string representing differences.
    missing → lines prefixed with 'Missing: ' (black)
    extra   → lines prefixed with 'Extra: ' (red if use_html else plain)
    If use_html True, extra lines are wrapped in <span style="color:red">.
    """
    lines = []
    for it in missing:
        lines.append(f"Missing: {it}")
    for it in extra:
        if use_html:
            lines.append(f'<span style="color:red">Extra: {it}</span>')
        else:
            lines.append(f"Extra: {it}")
    if not lines:
        return ""
    if use_html:
        return "<br>".join(lines)
    else:
        return "\n".join(lines)

# --- File upload ---
col1, col2 = st.columns(2)
with col1:
    std_file = st.file_uploader("Upload Standard Access Excel", type=["xlsx"], key="std")
with col2:
    clt_file = st.file_uploader("Upload Client Access Excel", type=["xlsx"], key="clt")

if not std_file or not clt_file:
    st.info("Upload both Standard and Client Excel files to see findings. Files must share the same headers.")
    st.stop()

# --- Read Excel files ---
try:
    std_df = pd.read_excel(std_file)
    clt_df = pd.read_excel(clt_file)
except Exception as e:
    st.error(f"Error reading Excel files: {e}")
    st.stop()

# Validate presence of SG column
if SG_COL not in std_df.columns or SG_COL not in clt_df.columns:
    st.error(f"Expected security group header `{SG_COL}` not found in one of the files. Found columns:\n- Standard: {list(std_df.columns)}\n- Client: {list(clt_df.columns)}")
    st.stop()

# Drop totally empty unnamed columns (common when exported)
std_df = std_df.loc[:, ~std_df.columns.str.contains('^Unnamed')]
clt_df = clt_df.loc[:, ~clt_df.columns.str.contains('^Unnamed')]

# Ensure both have same headers (except possible difference in whitespace)
std_cols = [c.strip() for c in std_df.columns]
clt_cols = [c.strip() for c in clt_df.columns]
if len(std_df.columns) != len(clt_df.columns) or any(a != b for a, b in zip(std_df.columns, clt_df.columns)):
    # They said headers are exact same; we warn but continue if the first column matches and other columns overlap
    st.warning("Column headers differ between files. Proceeding but please ensure both files use exactly the same headers.")
    
# --- Finding 1 & 2: security groups differences ---
std_sgs = set(std_df[SG_COL].dropna().astype(str).str.strip())
clt_sgs = set(clt_df[SG_COL].dropna().astype(str).str.strip())

only_in_std = sorted(list(std_sgs - clt_sgs))  # Security group that does not exist in tenant
only_in_clt = sorted(list(clt_sgs - std_sgs))  # Custom security group

# --- Finding 3: row-level comparisons for SGs in both ---
common_sgs = sorted(list(std_sgs & clt_sgs))

# columns to compare (all except SG_COL)
compare_columns = [c for c in std_df.columns if c != SG_COL]

# Build a differences DataFrame (plain text for export) and an HTML table for display
diff_rows = []
for sg in common_sgs:
    std_row = std_df[std_df[SG_COL].astype(str).str.strip() == sg].iloc[0]
    clt_row = clt_df[clt_df[SG_COL].astype(str).str.strip() == sg].iloc[0]
    row_diff = {SG_COL: sg}
    any_diff = False
    # For each column, compare
    for col in compare_columns:
        missing, extra = compare_cells(std_row.get(col, ""), clt_row.get(col, ""))
        # store plain text with newline prefixes for export
        row_diff[col] = build_difference_text(missing, extra, use_html=False)
        if missing or extra:
            any_diff = True
    if any_diff:
        diff_rows.append(row_diff)

diff_df = pd.DataFrame(diff_rows).fillna("")

# --- UI presentation ---
st.header("Findings")

c1, c2 = st.columns(2)
with c1:
    st.subheader("1) Security groups in Standard but NOT in Client")
    st.write("Label: **Security group that does not exist in tenant**")
    if only_in_std:
        st.write(f"Total: {len(only_in_std)}")
        st.write("\n".join(only_in_std))
    else:
        st.success("No security groups missing in tenant (Standard vs Client).")

with c2:
    st.subheader("2) Security groups in Client but NOT in Standard")
    st.write("Label: **Custom security group**")
    if only_in_clt:
        st.write(f"Total: {len(only_in_clt)}")
        st.write("\n".join(only_in_clt))
    else:
        st.success("No custom security groups found in Client.")

st.markdown("---")
st.subheader("3) Row-level differences (only rows with at least one difference)")

if diff_df.empty:
    st.success("No row-level differences found for security groups present in both files.")
else:
    st.write(f"Total security groups with differences: {len(diff_df)}")
    # Build HTML table for display: include red colored Extra lines
    display_df = diff_df.copy()
    for col in compare_columns:
        html_col = []
        for cell in display_df[col].astype(str).fillna(""):
            # reconstruct using compare_cells to get ordering
            # find missing and extra sets by parsing the cell's lines is unreliable, so recompute
            # We'll compute directly by locating the SG and column in original dfs:
            html_col.append("")  # placeholder; we'll fill below
        display_df[col] = html_col

    # Fill HTML cells properly by recomputing per SG
    html_rows = []
    for _, row in diff_df.iterrows():
        sg = row[SG_COL]
        std_row = std_df[std_df[SG_COL].astype(str).str.strip() == sg].iloc[0]
        clt_row = clt_df[clt_df[SG_COL].astype(str).str.strip() == sg].iloc[0]
        html_row = {SG_COL: sg}
        for col in compare_columns:
            missing, extra = compare_cells(std_row.get(col, ""), clt_row.get(col, ""))
            html_cell = build_difference_text(missing, extra, use_html=True)
            html_row[col] = html_cell
