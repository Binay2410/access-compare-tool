import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Security Group Access Comparator", layout="wide")
st.title("🔎 Security Group Access Comparator")
st.markdown("Upload **Standard** and **Client** Excel files (same headers). App shows 3 findings and lets you download results.")

# --- Constants
SG_COL = "Domains granted to Security Group"

# --- Helpers
def normalize_text_cell(cell):
    """Return list of stripped non-empty lines from a cell (handles NaN)."""
    if pd.isna(cell):
        return []
    # Convert numbers to string
    if isinstance(cell, (int, float)):
        cell = str(cell)
    # Split on newline (ALT+ENTER in Excel)
    parts = [p.strip() for p in str(cell).splitlines() if p.strip() != ""]
    return parts

def compare_cells(std_cell, clt_cell):
    """Return lists: missing (in std not in clt), extra (in clt not in std)."""
    sset = set(normalize_text_cell(std_cell))
    cset = set(normalize_text_cell(clt_cell))
    missing = sorted(list(sset - cset))
    extra = sorted(list(cset - sset))
    return missing, extra

def build_plaincell_text(missing, extra):
    """Return plain text cell (for export): Missing lines first, then Extra lines."""
    lines = []
    for it in missing:
        lines.append(f"Missing: {it}")
    for it in extra:
        lines.append(f"Extra: {it}")
    return "\n".join(lines)

def build_htmlcell(missing, extra):
    """Return HTML for display: Missing (black), Extra (red). Lines separated by <br>."""
    parts = []
    for it in missing:
        parts.append(f"Missing: {escape_html(it)}")
    for it in extra:
        parts.append(f'<span style="color:red">Extra: {escape_html(it)}</span>')
    return "<br>".join(parts) if parts else ""

def escape_html(s):
    """Minimal HTML escape to avoid tag injection inside cell text."""
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )

# --- File upload UI
st.header("Upload Client Access Excel")
clt_file = st.file_uploader("Upload Client Excel", type=["xlsx"])

# Load built-in Standard Excel stored in the repo
try:
    std_df = pd.read_excel("standard_data.xlsx")
except Exception as e:
    st.error("❌ Missing built-in file: standard_data.xlsx. Please add it to your GitHub repo.")
    st.stop()

if not clt_file:
    st.info("Please upload the Client Excel file to continue.")
    st.stop()

# Load Client Excel
try:
    clt_df = pd.read_excel(clt_file)
except Exception as e:
    st.error(f"Error reading Client Excel file: {e}")
    st.stop()


# Drop unnamed empty columns
std_df = std_df.loc[:, ~std_df.columns.str.contains('^Unnamed')]
clt_df = clt_df.loc[:, ~clt_df.columns.str.contains('^Unnamed')]

# Validate SG column present
if SG_COL not in std_df.columns or SG_COL not in clt_df.columns:
    st.error(f"Expected security group header `{SG_COL}` not found. Found columns:\n- Standard: {list(std_df.columns)}\n- Client: {list(clt_df.columns)}")
    st.stop()

# Normalize SG names (strip)
std_df[SG_COL] = std_df[SG_COL].astype(str).str.strip()
clt_df[SG_COL] = clt_df[SG_COL].astype(str).str.strip()

# Build sets for findings 1 & 2
std_sgs = set(std_df[SG_COL].dropna().astype(str).str.strip())
clt_sgs = set(clt_df[SG_COL].dropna().astype(str).str.strip())

only_in_std = sorted(list(std_sgs - clt_sgs))  # Security group that does not exist in tenant
only_in_clt = sorted(list(clt_sgs - std_sgs))  # Custom security group

# --- Finding 3: row-level differences for common SGs
common_sgs = sorted(list(std_sgs & clt_sgs))
compare_columns = [c for c in std_df.columns if c != SG_COL]

diff_rows_for_export = []  # list of dicts for excel export (plain text)
html_rows = []            # list of dicts with HTML cells for display

for sg in common_sgs:
    # get first matching rows (assumes single row per SG)
    std_row = std_df[std_df[SG_COL].astype(str).str.strip() == sg].iloc[0]
    clt_row = clt_df[clt_df[SG_COL].astype(str).str.strip() == sg].iloc[0]
    any_diff = False
    export_row = {SG_COL: sg}
    html_row = {SG_COL: sg}
    for col in compare_columns:
        missing, extra = compare_cells(std_row.get(col, ""), clt_row.get(col, ""))
        cell_plain = build_plaincell_text(missing, extra)
        cell_html = build_htmlcell(missing, extra)
        export_row[col] = cell_plain
        html_row[col] = cell_html
        if missing or extra:
            any_diff = True
    if any_diff:
        diff_rows_for_export.append(export_row)
        html_rows.append(html_row)

diff_export_df = pd.DataFrame(diff_rows_for_export).fillna("")
html_diff_df = pd.DataFrame(html_rows).fillna("")

# --- UI: Findings 1 & 2 (one per line)
st.header("Findings")

c1, c2 = st.columns(2)
with c1:
    st.subheader("1) Security groups that do not exist in tenant")
    st.markdown("**Description:** These security groups are available in Workday defaults based on areas implemented, but upon checking they were not found in the tenant.")
    st.write(f"Total: {len(only_in_std)}")
    if only_in_std:
        # print one per line
        st.text("\n".join(only_in_std))
    else:
        st.success("No security groups missing in tenant (Standard vs Client).")

with c2:
    st.subheader("2) Custom Security Groups")
    st.markdown("**Description:** These are the sustom security groups which might be created as part of requirement but they are highlighted here because they are not part of Workday defaults")
    st.write(f"Total: {len(only_in_clt)}")
    if only_in_clt:
        st.text("\n".join(only_in_clt))
    else:
        st.success("No custom security groups found in Client.")

st.markdown("---")

# --- UI: Finding 3 table (HTML)
st.subheader("3) Security Group difference report: ")

if html_diff_df.empty:
    st.success("No row-level differences found for security groups present in both files.")
else:
    st.write(f"Total security groups with differences: {len(html_diff_df)}")
    # Build HTML table (escape=False) so our <span style="color:red">Extra: ...</span> renders.
    # But we must ensure header order matches original compare_columns
    display_df = html_diff_df[[SG_COL] + compare_columns]
    html_table = display_df.to_html(index=False, escape=False)
    st.markdown(html_table, unsafe_allow_html=True)

# --- Download results as an Excel file (3 sheets)
def to_excel_bytes(missing_sgs, custom_sgs, differences_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame({SG_COL: missing_sgs}).to_excel(writer, sheet_name="Missing_SGs", index=False)
        pd.DataFrame({SG_COL: custom_sgs}).to_excel(writer, sheet_name="Custom_SGs", index=False)
        if differences_df.empty:
            pd.DataFrame().to_excel(writer, sheet_name="Row_Differences", index=False)
        else:
            differences_df.to_excel(writer, sheet_name="Row_Differences", index=False)
    return output.getvalue()


excel_bytes = to_excel_bytes(only_in_std, only_in_clt, diff_export_df)
now = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
st.download_button(
    label="Download findings as Excel",
    data=excel_bytes,
    file_name=f"security_findings_{now}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.info("Notes: 'Missing:' lines are plain black text. 'Extra:' lines are shown in red in the app UI (they are plain text in the downloaded Excel).")

