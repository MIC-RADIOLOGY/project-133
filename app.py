import streamlit as st
import pandas as pd
import openpyxl
import io
import math

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# =====================================================
# CLEAN TEXT HELPER
# =====================================================
def clean(text):
    if pd.isna(text):
        return ""
    return str(text).strip().upper()


# =====================================================
# LOAD & STRUCTURE THE CHARGE SHEET
# =====================================================
def load_charge_sheet(file):
    df = pd.read_excel(file, header=None)
    df.columns = ["EXAMINATION", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]

    structured = []

    current_category = None
    current_subcategory = None

    for _, row in df.iterrows():
        exam = clean(row["EXAMINATION"])

        # --- Detect MAIN CATEGORY (row starting with no tariff and bold style in sheet) ---
        if exam in [
            "ULTRA SOUND DOPPLERS",
            "ULTRA SOUND",
            "CT SCAN",
            "FLUROSCOPY",
            "X-RAY",
            "XRAY",
        ]:
            current_category = exam
            current_subcategory = None
            continue

        # --- Detect SUB CATEGORY (has no tariff, not TOTAL/FF, not category) ---
        if exam and pd.isna(row["TARIFF"]) and exam not in ["TOTAL", "FF", "CO-PAYMENT"]:
            current_subcategory = exam
            continue

        # --- Skip garbage rows ---
        if exam in ["FF", "TOTAL", "CO-PAYMENT", "", None]:
            continue

        # --- This is a REAL SCAN ROW ---
        t = pd.to_numeric(row["TARIFF"], errors="coerce")
        m = clean(row["MODIFIER"])

        qty_val = pd.to_numeric(row["QTY"], errors="coerce")
        qty = int(qty_val) if not math.isnan(qty_val) else 1

        amt_val = pd.to_numeric(row["AMOUNT"], errors="coerce")
        amt = float(amt_val) if not math.isnan(amt_val) else 0.0

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "TARIFF": t,
            "MODIFIER": m,
            "QTY": qty,
            "AMOUNT": amt
        })

    return pd.DataFrame(structured)


# =====================================================
# SAFE CELL WRITE
# =====================================================
def write_cell(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    try:
        cell.value = value
    except:
        for m in cell.parent.merged_cells.ranges:
            if cell.coordinate in m:
                top_left = m.coord.split(":")[0]
                ws[top_left].value = value
    return


# =====================================================
# WRITE TO QUOTATION TEMPLATE
# =====================================================
def fill_excel_template(template_file, patient, member, provider, scan_row):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    patient_cell = None
    member_cell = None
    provider_cell = None
    table_start_row = None
    desc_col = tarif_col = modi_col = qty_col = amt_col = None
    total_cell = None

    # --- Find template fields ---
    for row in ws.iter_rows():
        for c in row:
            if c.value:
                val = str(c.value).upper()

                if "PATIENT" in val and not patient_cell:
                    patient_cell = ws.cell(row=c.row, column=c.column + 1)

                elif "MEMBER" in val and not member_cell:
                    member_cell = ws.cell(row=c.row, column=c.column + 1)

                elif "EXAMINATION" in val or "PROVIDER" in val:
                    provider_cell = ws.cell(row=c.row, column=c.column + 1)

                elif "DESCRIPTION" in val:
                    table_start_row = c.row + 1
                    desc_col = c.column
                    tarif_col = c.column + 1
                    modi_col = c.column + 2
                    qty_col = c.column + 3
                    amt_col = c.column + 4

                elif val == "TOTAL":
                    total_cell = ws.cell(row=c.row, column=c.column + 6)

    # --- Write patient info ---
    if patient_cell: write_cell(ws, patient_cell.row, patient_cell.column, patient)
    if member_cell: write_cell(ws, member_cell.row, member_cell.column, member)
    if provider_cell: write_cell(ws, provider_cell.row, provider_cell.column, provider)

    # --- Write selected scan ---
    if table_start_row:
        write_cell(ws, table_start_row, desc_col, scan_row["SCAN"])
        write_cell(ws, table_start_row, tarif_col, scan_row["TARIFF"])
        write_cell(ws, table_start_row, modi_col, scan_row["MODIFIER"])
        write_cell(ws, table_start_row, qty_col, scan_row["QTY"])
        write_cell(ws, table_start_row, amt_col, scan_row["AMOUNT"])

    # --- Write total ---
    if total_cell:
        write_cell(ws, total_cell.row, total_cell.column, scan_row["AMOUNT"])

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# =====================================================
# STREAMLIT UI
# =====================================================
st.title("📄 Medical Quotation Generator")

if "df" not in st.session_state:
    st.session_state.df = None

charge_upload = st.file_uploader("Upload Charge Sheet", type=["xlsx"])
template_upload = st.file_uploader("Upload Quotation Template", type=["xlsx"])

# Patient info
patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if charge_upload and template_upload:
    if st.button("Load Charge Sheet"):
        st.session_state.df = load_charge_sheet(charge_upload)
        st.success("Charge sheet loaded successfully!")

    df = st.session_state.df

    if df is not None:
        # Main Category selection
        categories = sorted(df["CATEGORY"].dropna().unique())
        main_cat = st.selectbox("Select Main Category", categories)

        # Subcategory selection
        sub_df = df[df["CATEGORY"] == main_cat]
        subs = sorted(sub_df["SUBCATEGORY"].dropna().unique())
        subcat = st.selectbox("Select Subcategory", subs)

        # Scan selection
        scan_df = sub_df[sub_df["SUBCATEGORY"] == subcat]
        scan_names = scan_df["SCAN"].tolist()
        scan_name = st.selectbox("Select Scan", scan_names)

        scan_row = scan_df[scan_df["SCAN"] == scan_name].iloc[0]

        st.write("### Scan Details")
        st.write(scan_row)

        if st.button("Generate Quotation"):
            file = fill_excel_template(template_upload, patient, member, provider, scan_row)
            st.download_button(
                "Download Quotation",
                data=file,
                file_name=f"quotation_{patient}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("Upload both charge sheet and template to continue.")

