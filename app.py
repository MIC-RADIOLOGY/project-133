import streamlit as st
import pandas as pd
import openpyxl
import io
from copy import copy
from datetime import datetime

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config / heuristics ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN",
    "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND",
    "MRI"
}
GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# ---------- Helpers ----------
def clean_text(x) -> str:
    """Safely converts to string and cleans up common characters."""
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def u(x) -> str:
    """Cleans and converts text to uppercase."""
    return clean_text(x).upper()

def safe_int(x, default=1):
    """Safely converts a value to an integer."""
    try:
        x_str = str(x).replace(",", "").strip()
        return int(float(x_str))
    except Exception:
        return default

def safe_float(x, default=0.0):
    """Safely converts a value to a float, handling currency symbols and commas."""
    try:
        x_str = str(x).replace(",", "").replace("$", "").strip()
        return float(x_str)
    except Exception:
        return default

# ---------- Parser ----------
def load_charge_sheet(file) -> pd.DataFrame:
    """Parses the raw Excel charge sheet into a structured DataFrame."""
    df_raw = pd.read_excel(file, header=None, dtype=object)

    while df_raw.shape[1] < 5:
        df_raw[df_raw.shape[1]] = None
    df_raw = df_raw.iloc[:, :5]
    df_raw.columns = ["A_EXAM", "B_TARIFF", "C_MOD", "D_QTY", "E_AMOUNT"]

    structured = []
    current_category = None
    current_subcategory = None

    for idx, r in df_raw.iterrows():
        exam = clean_text(r["A_EXAM"])
        if exam == "":
            continue
        exam_u = exam.upper()

        # Check if all non-essential columns are blank (heuristic for header/category rows)
        is_header_row = (
            (pd.isna(r["B_TARIFF"]) or clean_text(r["B_TARIFF"]) == "") and
            (pd.isna(r["C_MOD"]) or clean_text(r["C_MOD"]) == "") and
            (pd.isna(r["D_QTY"]) or clean_text(r["D_QTY"]) == "") and
            (pd.isna(r["E_AMOUNT"]) or clean_text(r["E_AMOUNT"]) == "")
        )

        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        if is_header_row:
            # If it looks like a header but isn't a main category, treat as subcategory
            current_subcategory = exam
            continue

        if exam_u in GARBAGE_KEYS:
            continue

        # If we reach here, treat it as a scan/procedure item
        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "TARIFF": safe_float(r["B_TARIFF"], None),
            "MODIFIER": clean_text(r["C_MOD"]),
            "QTY": safe_int(r["D_QTY"], 1),
            "AMOUNT": safe_float(r["E_AMOUNT"], 0.0)
        })

    return pd.DataFrame(structured)

# ---------- Excel Template Helpers ----------
def write_safe(ws, r, c, value):
    """Writes value to cell, handling merged cells and bounds safely."""
    if c is None:
        return
    try:
        cell = ws.cell(row=r, column=c)
    except Exception:
        return
    try:
        cell.value = value
    except Exception:
        # Fallback for merged cells
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                topcell = mr.coord.split(":")[0]
                ws[topcell].value = value
                return

def find_template_positions(ws):
    """Dynamically finds key column and row positions in the Excel template."""
    pos = {}
    
    # Standardize column names for internal use
    header_map = {
        "DESCRIPTION": ["DESCRIPTION", "PROCEDURE", "EXAMINATION", "TEST NAME"],
        "RATE": ["TARIFF", "TARRIF", "RATE", "PRICE"],
        "MOD": ["MOD", "MODIFIER"],
        "QTY": ["QTY", "QUANTITY", "NO", "NUMBER"],
        "LINE_TOTAL": ["FEES", "CHARGE", "AMOUNT PER ITEM"],
        "GRAND_TOTAL": ["AMOUNT", "TOTAL", "LINE TOTAL", "TOTAL AMOUNT", "TOTAL FEES"]
    }

    found_headers = {key: None for key in header_map}
    header_row_candidate = None

    for row in ws.iter_rows(min_row=1, max_row=300):
        for cell in row:
            if not cell.value:
                continue
            cell_text = str(cell.value).upper().strip()
            
            # Detect patient, member, provider, date cells
            if "PATIENT" in cell_text and "patient_cell" not in pos:
                pos["patient_cell"] = (cell.row, cell.column)
            if "MEMBER" in cell_text and "member_cell" not in pos:
                pos["member_cell"] = (cell.row, cell.column)
            if ("PROVIDER" in cell_text or "EXAMINATION" in cell_text) and "provider_cell" not in pos:
                pos["provider_cell"] = (cell.row, cell.column)
            if "DATE" in cell_text and "date_cell" not in pos:
                pos["date_cell"] = (cell.row, cell.column)
            
            # Detect headers
            for key, variants in header_map.items():
                for v in variants:
                    if v.upper() in cell_text:
                        found_headers[key] = cell.column

        # Heuristic: If we find at least three required columns (DESC, RATE, QTY) in this row, 
        # assume it's the header row for the data block.
        required_keys = ["DESCRIPTION", "RATE", "QTY"]
        if all(found_headers[k] for k in required_keys) and header_row_candidate is None:
            header_row_candidate = row[0].row
            
    pos["cols"] = {k: v for k, v in found_headers.items() if v is not None}
    
    # The data should start one row after the detected header row. Fallback to 22.
    pos["data_start_row"] = (header_row_candidate + 1) if header_row_candidate else 22

    # CRITICAL CHECK: Raise an error if required columns for data entry are missing
    required = ["DESCRIPTION", "RATE", "MOD", "QTY", "LINE_TOTAL"]
    missing = [col for col in required if col not in pos["cols"]]
    if missing:
        raise ValueError(f"Your quotation template is missing these required data columns: {', '.join(missing)}. Please ensure the headers are visible and correctly spelled.")

    return pos

def replace_after_colon_in_same_cell(ws, row, col, new_value):
    """Updates cell value, preserving text before a colon, if present."""
    cell = ws.cell(row=row, column=col)
    # Check for merged cell and get the top-left cell if merged
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            cell = ws[mr.coord.split(":")[0]]
            break
            
    old = str(cell.value) if cell.value else ""
    if ":" in old:
        left = old.split(":", 1)[0]
        cell.value = f"{left}: {new_value}"
    else:
        cell.value = new_value

# The function write_value_preserve_borders is available but not strictly needed 
# if write_safe handles merged cells correctly.

# ---------- Fill Template ----------
def fill_excel_template(template_file, patient, member, provider, scan_rows):
    """Fills the Excel template with patient data and scan details."""
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    
    # This call includes robust error checking
    pos = find_template_positions(ws) 

    # Fill patient info
    if "patient_cell" in pos:
        r, c = pos["patient_cell"]
        replace_after_colon_in_same_cell(ws, r, c, patient)
    if "member_cell" in pos:
        r, c = pos["member_cell"]
        replace_after_colon_in_same_cell(ws, r, c, member)
    if "provider_cell" in pos:
        r, c = pos["provider_cell"]
        replace_after_colon_in_same_cell(ws, r, c, provider)
    if "date_cell" in pos:
        r, c = pos["date_cell"]
        today_str = datetime.now().strftime("%d/%m/%Y")
        # Write date value below the 'DATE' header cell
        write_safe(ws, r + 1, c, today_str)

    # Write scan items
    if "cols" in pos:
        start_row = pos.get("data_start_row", 22) # Use dynamic start row
        cols = pos["cols"]

        # Write each scan on a single row
        for idx, sr in enumerate(scan_rows):
            rowptr = start_row + idx
            
            # Use get() with None/default fallback for safety
            scan = sr.get("SCAN")
            tariff = sr.get("TARIFF")
            modifier = sr.get("MODIFIER")
            qty = sr.get("QTY")
            amount = sr.get("AMOUNT")

            # Write values using the dynamically found column indices and ensuring numerical fields are safe
            write_safe(ws, rowptr, cols.get("DESCRIPTION"), scan)
            write_safe(ws, rowptr, cols.get("RATE"), tariff if tariff is not None else 0.0) 
            write_safe(ws, rowptr, cols.get("MOD"), modifier or "")
            write_safe(ws, rowptr, cols.get("QTY"), qty if qty is not None else 1)
            write_safe(ws, rowptr, cols.get("LINE_TOTAL"), amount if amount is not None else 0.0) 

        # Calculate and write the total amount
        total_amt = sum([safe_float(r.get("AMOUNT", 0.0), 0.0) for r in scan_rows])
        
        grand_total_col = cols.get("GRAND_TOTAL")
        
        if grand_total_col:
            # Write the total two rows after the last written scan, using the GRAND_TOTAL column
            total_row_ptr = start_row + len(scan_rows) + 2 
            write_safe(ws, total_row_ptr, grand_total_col, total_amt)
        else:
            # Fallback for templates without a detectable GRAND_TOTAL header
            write_safe(ws, 22, 7, total_amt) # Column 7 = G

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------- Streamlit UI Logic Functions ----------
def parse_and_set_state(uploaded_charge):
    """Handles the parsing logic for the uploaded charge sheet."""
    try:
        parsed = load_charge_sheet(uploaded_charge)
        st.session_state.parsed_df = parsed
        st.session_state.selected_rows = []
        st.success("Charge sheet parsed successfully.")
    except Exception as e:
        st.error(f"Failed to parse charge sheet: {e}")
        st.session_state.parsed_df = None

def display_data_selection(df):
    """Displays category selectors and filters the data."""
    st.subheader("4. Data Selection")
    cats = [c for c in sorted(df["CATEGORY"].dropna().unique())]

    if not cats:
        st.warning("No categories found in the uploaded charge sheet.")
        st.session_state.selected_rows = []
        return

    col1, col2 = st.columns([3, 4])
    with col1:
        main_sel = st.selectbox("Select Main Category", ["-- choose --"] + cats, key="main_select")

    if main_sel == "-- choose --":
        st.session_state.selected_rows = []
        return

    # Subcategory selection
    subs = [s for s in sorted(df[df["CATEGORY"] == main_sel]["SUBCATEGORY"].dropna().unique())]
    sub_sel = "-- all --"

    if subs:
        sub_sel = st.selectbox("Select Subcategory", ["-- all --"] + subs, key="sub_select")

    # Filtering logic
    if sub_sel == "-- all --":
        scans_for_cat = df[df["CATEGORY"] == main_sel].reset_index(drop=True)
    else:
        scans_for_cat = df[(df["CATEGORY"] == main_sel) & (df["SUBCATEGORY"] == sub_sel)].reset_index(drop=True)

    st.session_state.selected_rows = scans_for_cat.to_dict(orient="records")

def display_results_and_download_ui(uploaded_template, patient, member, provider):
    """Displays selected scans and handles the final generation/download."""
    st.markdown("---")
    st.subheader("5. Selected Scans and Quotation Generation")

    if not st.session_state.get("selected_rows"):
        st.info("No scans found for the current category/subcategory or none selected.")
        st.dataframe(pd.DataFrame(columns=["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]))
        return

    sel_df = pd.DataFrame(st.session_state.selected_rows)
    display_df = sel_df[["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]]
    st.dataframe(display_df.reset_index(drop=True))

    total_amt = sum([safe_float(r.get("AMOUNT", 0.0), 0.0) for r in st.session_state.selected_rows])
    st.markdown(f"**Total Quotation Amount:** ${total_amt:,.2f}")

    if uploaded_template:
        if st.button("Generate Quotation and Download Excel"):
            try:
                out = fill_excel_template(uploaded_template, patient, member, provider, st.session_state.selected_rows)
                st.download_button(
                    "Download Quotation",
                    data=out,
                    file_name=f"quotation_{(patient or 'patient').replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except ValueError as ve:
                st.error(f"Template Error: {ve}")
            except Exception as e:
                st.error(f"Failed to generate quotation: {e}")
    else:
        st.warning("Upload a Quotation Template to enable generation.")


# ---------- Streamlit UI Entry Point ----------
st.title("🏥 Medical Quotation Generator — Full Tariff Capture")
st.info("Follow the steps below to generate a quotation from a raw charge sheet and a template.")

# Input Section
debug_mode = st.checkbox("Show parsing debug output", value=False)
uploaded_charge = st.file_uploader("1. Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("2. Upload Quotation Template (Excel)", type=["xlsx"])

st.subheader("3. Patient and Provider Information")
col_p, col_m, col_pr = st.columns(3)
with col_p:
    patient = st.text_input("Patient Name", key="patient")
with col_m:
    member = st.text_input("Medical Aid / Member Number", key="member")
with col_pr:
    provider = st.text_input("Medical Aid Provider", value="CIMAS", key="provider")


# Core Logic Flow
if uploaded_charge:
    if st.button("Parse Charge Sheet", key="parse_button"):
        parse_and_set_state(uploaded_charge)

    if "parsed_df" in st.session_state and st.session_state.parsed_df is not None:
        df = st.session_state.parsed_df

        if debug_mode:
            st.subheader("📊 Debug Output: Raw Parsed Data")
            st.write("Parsed DataFrame columns:", df.columns.tolist())
            st.dataframe(df.head(200))
        
        # Display selection UI and update st.session_state.selected_rows
        display_data_selection(df)

        # Display results and handle download
        display_results_and_download_ui(uploaded_template, patient, member, provider)

else:
    st.info("Upload a charge sheet to begin.")
