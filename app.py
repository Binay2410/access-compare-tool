# app.py
import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

st.set_page_config(page_title="Security Group Access Comparator (MVP+)", layout="wide")
st.title("🔎 Security Group Access Comparator (MVP+)")

# --- Constants: exact header names as you provided
HEADERS = [
    "Domains granted to Security Group",
    "Reports and Tasks - Modify Access",
    "Reports and Tasks - View Access",
    "Integrations - Put Access",
    "Integrations - Get Access",
    "Business Process Types granted to Security Group - Initiate Access",
    "Business Process Types granted to Security Group - Enrichment Access",
    "Business Process Types granted to Security Group - Approve Access",
    "Business Process Types granted to Security Group - View Access",
    "Business Process Types granted to Security Group - View Completed Access"
]

SG_COL = HEADERS[0]

# Filenames for built-in and output files
STANDARD_FILE = "standard_data.xlsx"
ACCEPTED_RISK_FILE = "accepted_risk.xlsx"
UPDATE_WORKDAY_FILE = "update_workday.xlsx"

REQUIRED_FILES = [STANDARD_FILE]

# --- Helpers
def ensure_file_with_headers(path, headers):
    """Create Excel file with header row if it doesn't exist."""
    if not os.path.exists(path):
        df = pd.DataFrame(columns=headers)
        df.to_excel(path, index=False)
        st.info(f"Created missing file: {path}")

def read_excel_safe(path_or_file):
    """Read either a path (str) or an uploaded file-like object."""
    try:
        if isinstance(path_or_file, str):
            return pd.read_excel(path_or_file)
        else:
            return pd.read_excel(path_or_file)
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        raise

def normalize_text_cell(cell):
    """Return list of stripped non-empty lines from a cell (handles NaN)."""
    if pd.isna(cell):
        return []
    if isinstance(cell, (int, float)):
        cell = str(cell)
    parts = [p.strip() for p in str(cell).splitlines() if p.strip() != ""]
    return parts

def compare_cells(std_cell, clt_cell, ignore_set=None):
    """
    Return missing (in std not in clt) and extra (in clt not in std),
    after removing any items in ignore_set (set of items to ignore).
    """
    sset = set(normalize_text_cell(std_cell))
    cset = set(normalize_text_cell(clt_cell))
    if ignore_set:
        sset = set([x for x in sset if x not in ignore_set])
        cset = set([x for x in cset if x not in ignore_set])
    missing = sorted(list(sset - cset))
    extra = sorted(list(cset - sset))
    return missing, extra

def build_plaincell_text(missing, extra):
    """Return plain text lines (Missing then Extra) for export."""
    lines = []
    for it in missing:
        lines.append(f"Missing: {it}")
    for it in extra:
        lines.append(f"Extra: {it}")
    return "\n".join(lines)

def html_escape(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def build_htmlcell(missing, extra):
    """Return HTML with Missing lines black, Extra lines red."""
    parts = []
    for it in missing:
        parts.append(f"Missing: {html_escape(it)}")
    for it in extra:
        parts.append(f'<span style="color:red">Extra: {html_escape(it)}</span>')
    return "<br>".join(parts) if parts else ""

def load_accept_update_sets():
    """
    Load accepted_risk.xlsx and update_workday.xlsx and return a dict mapping:
    ignore_map[(sg, column)] = set(items)  (items that should be ignored in differences)
    """
    ignore_map = {}
    # Ensure files exist
    ensure_file_with_headers(ACCEPTED_RISK_FILE, HEADERS)
    ensure_file_with_headers(UPDATE_WORKDAY_FILE, HEADERS)

    # Load both
    for path in [ACCEPTED_RISK_FILE, UPDATE_WORKDAY_FILE]:
        try:
            df = pd.read_excel(path)
        except Exception:
            df = pd.DataFrame(columns=HEADERS)
        if SG_COL not in df.columns:
            continue
        for _, r in df.iterrows():
            sg = str(r.get(SG_COL, "")).strip()
            if not sg:
                continue
            for col in HEADERS[1:]:
                cell = r.get(col, "")
                for item in normalize_text_cell(cell):
                    ignore_map.setdefault((sg, col), set()).add(item)
    return ignore_map

def append_raw_item_to_excel(path, sg, column, item):
    """
    Append the raw item (no prefix) to the Excel at `path`, to the row matching sg,
    in the column `column`. If the row doesn't exist, create it.
    Append as new line inside the cell, preserving existing content.
    """
    ensure_file_with_headers(path, HEADERS)
    df = pd.read_excel(path)
    # Ensure all headers exist in the df
    for h in HEADERS:
        if h not in df.columns:
            df[h] = ""

    # Find row index for sg (first occurrence)
    mask = df[SG_COL].astype(str).str.strip() == str(sg).strip()
    if mask.any():
        idx = df[mask].index[0]
        existing = df.at[idx, column]
        existing_lines = normalize_text_cell(existing)
        # Avoid duplicate appended value if already present
        if item in existing_lines:
            return False  # nothing appended
        # Append
        if pd.isna(existing) or str(existing).strip() == "":
            newcell = item
        else:
            newcell = str(existing) + "\n" + item
        df.at[idx, column] = newcell
    else:
        # create new row with empty columns, set SG and the item in column
        new = {h: "" for h in HEADERS}
        new[SG_COL] = sg
        new[column] = item
        df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
    # Save back
    df.to_excel(path, index=False)
    return True

def to_excel_bytes(missing_sgs, custom_sgs, differences_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame({SG_COL: missing_sgs}).to_excel(writer, sheet_name="Missing_SGs", index=False)
        pd.DataFrame({SG_COL: custom_sgs}).to_excel(writer, sheet_name="Custom_SGs", index=False)
        if differences_df is None or differences_df.empty:
            pd.DataFrame(columns=HEADERS).to_excel(writer, sheet_name="Row_Differences", index=False)
        else:
            differences_df.to_excel(writer, sheet_name="Row_Differences", index=False)
    return output.getvalue()

# --- Ensure standard file exists
for f in REQUIRED_FILES:
    if not os.path.exists(f):
        st.error(f"Required file `{f}` not found in app repo. Please add it and redeploy.")
        st.stop()

# Ensure accepted/update files exist (they will be created on demand)
ensure_file_with_headers(ACCEPTED_RISK_FILE, HEADERS)
ensure_file_with_headers(UPDATE_WORKDAY_FILE, HEADERS)

# --- Simple page router using query params
query_params = st.experimental_get_query_params()
page = query_params.get("page", ["summary"])[0]
selected_sg_for_details = query_params.get("sg", [""])[0]

# --- Upload Client file (only client)
st.sidebar.header("Upload / Actions")
uploaded_client = st.sidebar.file_uploader("Upload Client Access Excel (single file)", type=["xlsx"])

if st.sidebar.button("Reset to Summary View"):
    st.experimental_set_query_params(page="summary")
    st.experimental_rerun()

# --- Read standard baseline
try:
    std_df = read_excel_safe(STANDARD_FILE)
except Exception:
    st.error("Unable to load standard baseline file.")
    st.stop()

# Validate standard headers
if SG_COL not in std_df.columns:
    st.error(f"Standard file missing required column `{SG_COL}`.")
    st.stop()

# If client file not uploaded yet, prompt and stop on summary page
if uploaded_client is None and page == "summary":
    st.info("Please upload the Client Excel file from the left sidebar to run analysis.")
    st.stop()

# When client uploaded, read it
clt_df = None
if uploaded_client:
    try:
        clt_df = read_excel_safe(uploaded_client)
    except Exception as e:
        st.error(f"Error reading uploaded client file: {e}")
        st.stop()

    # Validate client headers
    if SG_COL not in clt_df.columns:
        st.error(f"Uploaded client file missing required column `{SG_COL}`.")
        st.stop()

# --- Load ignore map from accepted_risk and update_workday
ignore_map = load_accept_update_sets()

# --- MAIN: summary page
if page == "summary":
    st.header("Findings")

    # Normalize and drop unnamed columns from both
    std_df = std_df.loc[:, std_df.columns.notna()]
    std_df = std_df[[c for c in std_df.columns if not str(c).startswith("Unnamed")]]
    clt_df = clt_df.loc[:, clt_df.columns.notna()]
    clt_df = clt_df[[c for c in clt_df.columns if not str(c).startswith("Unnamed")]]

    # Normalize SG text
    std_df[SG_COL] = std_df[SG_COL].astype(str).str.strip()
    clt_df[SG_COL] = clt_df[SG_COL].astype(str).str.strip()

    std_sgs = set(std_df[SG_COL].dropna().astype(str).str.strip())
    clt_sgs = set(clt_df[SG_COL].dropna().astype(str).str.strip())

    only_in_std = sorted(list(std_sgs - clt_sgs))
    only_in_clt = sorted(list(clt_sgs - std_sgs))
    common_sgs = sorted(list(std_sgs & clt_sgs))

    # Finding 1 & 2 display with one-per-line
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1) Security groups in Standard but NOT in Client")
        st.markdown("**Label:** Security group that does not exist in tenant")
        st.write(f"Total: {len(only_in_std)}")
        if only_in_std:
            st.text("\n".join(only_in_std))
        else:
            st.success("No security groups missing in tenant (Standard vs Client).")
    with c2:
        st.subheader("2) Security groups in Client but NOT in Standard")
        st.markdown("**Label:** Custom security group")
        st.write(f"Total: {len(only_in_clt)}")
        if only_in_clt:
            st.text("\n".join(only_in_clt))
        else:
            st.success("No custom security groups found in Client.")

    st.markdown("---")

    # Build row-level differences for common SGs, excluding items present in accepted/update files
    compare_columns = [c for c in std_df.columns if c != SG_COL]
    diff_rows_export = []
    diff_rows_html = []

    for sg in common_sgs:
        std_row = std_df[std_df[SG_COL].astype(str).str.strip() == sg].iloc[0]
        clt_row = clt_df[clt_df[SG_COL].astype(str).str.strip() == sg].iloc[0]
        any_diff = False
        export_row = {SG_COL: sg}
        html_row = {SG_COL: sg}
        for col in compare_columns:
            # build ignore set for this sg+col
            ignore_set = ignore_map.get((sg, col), set())
            missing, extra = compare_cells(std_row.get(col, ""), clt_row.get(col, ""), ignore_set=ignore_set)
            export_text = build_plaincell_text(missing, extra)
            html_text = build_htmlcell(missing, extra)
            export_row[col] = export_text
            html_row[col] = html_text
            if missing or extra:
                any_diff = True
        if any_diff:
            diff_rows_export.append(export_row)
            diff_rows_html.append(html_row)

    diff_export_df = pd.DataFrame(diff_rows_export).fillna("")
    diff_html_df = pd.DataFrame(diff_rows_html).fillna("")

    st.subheader("3) Row-level differences (only rows with at least one difference)")
    if diff_html_df.empty:
        st.success("No row-level differences found for security groups present in both files.")
    else:
        st.write(f"Total security groups with differences: {len(diff_html_df)}")
        # Add a Details button per row; we will render the HTML table and add buttons below it per row.
        display_df = diff_html_df[[SG_COL] + compare_columns].copy()
        html_table = display_df.to_html(index=False, escape=False)
        st.markdown(html_table, unsafe_allow_html=True)

        st.markdown("**Actions:** Click Details for a security group to view per-element actions.")
        for idx, row in diff_export_df.iterrows():
            sg = row[SG_COL]
            btn_key = f"details_{sg}_{idx}"
            if st.button(f"Details: {sg}", key=btn_key):
                # navigate to details page with sg
                st.experimental_set_query_params(page="details", sg=sg)
                st.experimental_rerun()

    # Download findings
    excel_bytes = to_excel_bytes(only_in_std, only_in_clt, diff_export_df)
    now = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    st.download_button(
        label="Download findings as Excel",
        data=excel_bytes,
        file_name=f"security_findings_{now}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- DETAILS page for a single Security Group
elif page == "details":
    sg = selected_sg_for_details
    if not sg:
        st.error("No security group selected.")
        st.stop()
    st.header(f"Details — {sg}")
    st.markdown("Per-element actions. For each element you can choose `Accept Risk` (⚠️) or `Update Workday` (🔧). Raw item will be appended to the respective file and will be ignored in future analyses.")

    # Reload standard and client and ignore map to ensure up-to-date
    std_df = read_excel_safe(STANDARD_FILE)
    clt_df = read_excel_safe(uploaded_client) if uploaded_client else pd.DataFrame(columns=HEADERS)
    ignore_map = load_accept_update_sets()

    # fetch rows
    std_row = std_df[std_df[SG_COL].astype(str).str.strip() == sg]
    clt_row = clt_df[clt_df[SG_COL].astype(str).str.strip() == sg]

    if std_row.empty or clt_row.empty:
        st.warning("Security Group is not present in both Standard and Client. Details page expects SG present in both.")
        st.write("If you still want to act on items for this SG, upload a client file where this SG is present.")
        if st.button("Back to Summary"):
            st.experimental_set_query_params(page="summary")
            st.experimental_rerun()
        st.stop()

    std_row = std_row.iloc[0]
    clt_row = clt_row.iloc[0]

    compare_columns = [c for c in std_df.columns if c != SG_COL]

    # show per-column items
    for col in compare_columns:
        st.subheader(col)
        ignore_set = ignore_map.get((sg, col), set())
        missing, extra = compare_cells(std_row.get(col, ""), clt_row.get(col, ""), ignore_set=ignore_set)
        if not missing and not extra:
            st.markdown("_No actionable differences in this column after applying accepted/update filters._")
            continue

        # Display Missing items (black text) with buttons
        if missing:
            st.markdown("**Missing (in Standard but not in Client)**")
            for it in missing:
                st.write(it)
                col_a, col_b = st.columns([1,1])
                # unique keys for buttons
                btn_accept_key = f"accept_risk__{sg}__{col}__M__{it}"
                btn_update_key = f"update_workday__{sg}__{col}__M__{it}"
                if col_a.button(f"⚠️ Accept Risk", key=btn_accept_key):
                    appended = append_raw_item_to_excel(ACCEPTED_RISK_FILE, sg, col, it)
                    if appended:
                        st.success(f"Appended to {ACCEPTED_RISK_FILE}")
                    else:
                        st.info("Item already exists in the target file; no change.")
                    # refresh by redirecting back to details (so ignore_map updates)
                    st.experimental_set_query_params(page="details", sg=sg)
                    st.experimental_rerun()
                if col_b.button(f"🔧 Update Workday", key=btn_update_key):
                    appended = append_raw_item_to_excel(UPDATE_WORKDAY_FILE, sg, col, it)
                    if appended:
                        st.success(f"Appended to {UPDATE_WORKDAY_FILE}")
                    else:
                        st.info("Item already exists in the target file; no change.")
                    st.experimental_set_query_params(page="details", sg=sg)
                    st.experimental_rerun()

        # Display Extra items (red) with buttons
        if extra:
            st.markdown("**Extra (in Client but not Standard)**")
            for it in extra:
                # show red text using markdown span
                st.markdown(f'<span style="color:red">{html_escape(it)}</span>', unsafe_allow_html=True)
                col_a, col_b = st.columns([1,1])
                btn_accept_key = f"accept_risk__{sg}__{col}__E__{it}"
                btn_update_key = f"update_workday__{sg}__{col}__E__{it}"
                if col_a.button(f"⚠️ Accept Risk", key=btn_accept_key):
                    appended = append_raw_item_to_excel(ACCEPTED_RISK_FILE, sg, col, it)
                    if appended:
                        st.success(f"Appended to {ACCEPTED_RISK_FILE}")
                    else:
                        st.info("Item already exists in the target file; no change.")
                    st.experimental_set_query_params(page="details", sg=sg)
                    st.experimental_rerun()
                if col_b.button(f"🔧 Update Workday", key=btn_update_key):
                    appended = append_raw_item_to_excel(UPDATE_WORKDAY_FILE, sg, col, it)
                    if appended:
                        st.success(f"Appended to {UPDATE_WORKDAY_FILE}")
                    else:
                        st.info("Item already exists in the target file; no change.")
                    st.experimental_set_query_params(page="details", sg=sg)
                    st.experimental_rerun()

    if st.button("Back to Summary"):
        st.experimental_set_query_params(page="summary")
        st.experimental_rerun()

else:
    st.error("Unknown page.")
