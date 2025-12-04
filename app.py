import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

MAIN_CATEGORIES = [
    "ULTRASOUND",
    "ULTRASOUND DOPPLERS",
    "CT SCAN",
    "FLUROSCOPY",
    "X-RAY",
]

INVALID_WORDS = ["TOTAL", "CO", "PAYMENT", "FF"]

def clean_text(x):
    if pd.isna(x): return ""
    return str(x).strip().upper()

# ----------------------------------------------------
# LOAD & STRUCTURE CHARGE SHEET
# ----------------------------------------------------
def load_charge_sheet(file):
    df = pd.read_excel(file, header=None)
    df.columns = ["EXAM", "TARIFF", "MOD", "QTY", "AMOUNT"]

    df["EXAM"] = df["EXAM"].astype(str)

    # Remove rows like TOTAL, CO-PAYMENT, FF etc.
    df = df[~df["EXAM"].str.upper().apply(lambda x: any(w in x for w in INVALID_WORDS))]

    # Build CATEGORY → SUBCATEGORY → SCAN
    structured = []
    current_cat = None
    current_sub = None

    for _, row in df.iterrows():
        exam = clean_text(row["EXAM"])

        # Is a main category?
        if exam in MAIN_CATEGORIES:
            current_cat = exam
            current_sub = None
            continue

        # Is a subcategory (usually bold)?
        if exam.isupper() and len(exam.split()) <= 3:
            current_sub = exam
            continue

        # Otherwise → this is the actual scan
        if current_cat:
            structured.append({
                "CATEGORY": current_cat,
                "SUBCATEGORY": current_sub,
                "SCAN": exam,
                "TARIFF": str(row["TARIFF"]),
                "MODIFIER": str(row["MOD"]),
                "QTY": int(pd.to_numeric(row["QTY"], errors="coerce") or 1),
                "AMOUNT": float(pd.to_numeric(row["AMOUNT"], errors="coerce") or 0)
            })

    return pd.DataFrame(structured)

# ----------------------------------------------------
# FILL TEMPLATE
# ----------------------------------------------------
def set_cell_value(cell, value):
    try:
        cell.value = value
    except:
        # handle merged cells
        for m in cell.parent.merged_cells.ranges:
            if cell.coordinate in m:
                top_left = m.coord.split(":")[0]
                cell.parent[top_left].value = value

def fill_excel_template(template_file, patient, member, provider, scan_row):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    patient_cell = member_cell = provider_cell = None
    scan_start = None
    total_cell = None

    for row in ws.iter_rows():
        for c in row:
            if c.value:
                text = str(c.value).upper()
                if "PATIENT" in text:
                    patient_cell = ws.cell(c.row, c.column+1)
                if "MEMBER" in text:
                    member_cell = ws.cell(c.row, c.column+1)
                if "PROVIDER" in text or "EXAMINATION" in text:
                    provider_cell = ws.cell(c.row, c.column+1)
                if "DESCRIPTION" in text:
                    scan_start = c.row + 1
                    desc_col = c.column
                if "TOTAL" in text:
                    total_cell = ws.cell(c.row, c.column+6)

    # fill patient info
    if patient_cell: set_cell_value(patient_cell, patient)
    if member_cell: set_cell_value(member_cell, member)
    if provider_cell: set_cell_value(provider_cell, provider)

    # fill scan row
    if scan_start:
        set_cell_value(ws.cell(scan_start, desc_col), scan_row["SCAN"])
        set_cell_value(ws.cell(scan_start, desc_col+1), scan_row["TARIFF"])
        set_cell_value(ws.cell(scan_start, desc_col+2), scan_row["MODIFIER"])
        set_cell_value(ws.cell(scan_start, desc_col+3), scan_row["QTY"])
        set_cell_value(ws.cell(scan_start, desc_col+4), scan_row["AMOUNT"])

    if total_cell:
        set_cell_value(total_cell, scan_row["AMOUNT"])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ----------------------------------------------------
# STREAMLIT UI
# ----------------------------------------------------
st.title("📄 Medical Quotation Generator (Corrected Version)")

uploaded_charge = st.file_uploader("Upload Charge Sheet", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid Number")
provider = st.text_input("Provider (CIMAS)", value="CIMAS")

if uploaded_charge and uploaded_template:
    if st.button("Load Charge Sheet"):
        df = load_charge_sheet(uploaded_charge)
        st.session_state.df = df
        st.success("Charge sheet loaded successfully!")

    if "df" in st.session_state:
        df = st.session_state.df

        # Step 1 – Select main category
        cat = st.selectbox("Select Main Category", df["CATEGORY"].unique())

        # Step 2 – select subcategory
        sub_list = df[df["CATEGORY"] == cat]["SUBCATEGORY"].unique()
        sub = st.selectbox("Select Sub-Category", sub_list)

        # Step 3 – select scan
        scans = df[(df["CATEGORY"] == cat) & (df["SUBCATEGORY"] == sub)]
        scan_name = st.selectbox("Select Scan", scans["SCAN"].unique())

        scan_row = scans[scans["SCAN"] == scan_name].iloc[0]

        st.write("### Scan Details")
        st.write(scan_row)

        if st.button("Generate Quotation"):
            output = fill_excel_template(
                uploaded_template, patient, member, provider, scan_row
            )
            st.download_button(
                "Download Quotation",
                data=output,
                file_name=f"quotation_{patient}.xlsx"
            )
