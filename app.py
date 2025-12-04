import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# -----------------------------------
# LOAD CHARGE SHEET
# -----------------------------------
def load_charge_sheet(file):
    df = pd.read_excel(file)
    st.write("DEBUG: Columns detected in charge sheet:", df.columns.tolist())  # Debug print
    
    df.columns = [c.strip().upper() for c in df.columns]
    st.write("DEBUG: Columns after cleaning:", df.columns.tolist())  # Debug print
    
    if "TARIFF" not in df.columns:
        st.error("Column 'TARIFF' not found in charge sheet. Please check your Excel file.")
        return None
    
    df = df[df["TARIFF"].notna()]
    df["EXAMINATION"] = df["EXAMINATION"].astype(str)
    df["TARIFF"] = df["TARIFF"].astype(str)
    df["MODIFIER"] = df["MODIFIER"].astype(str)
    
    # Make sure QUANTITY and AMOUNT columns exist and have correct types
    if "QUANTITY" not in df.columns or "AMOUNT" not in df.columns:
        st.error("Charge sheet must contain 'QUANTITY' and 'AMOUNT' columns.")
        return None
    
    df["QUANTITY"] = pd.to_numeric(df["QUANTITY"], errors='coerce').fillna(0).astype(int)
    df["AMOUNT"] = pd.to_numeric(df["AMOUNT"], errors='coerce').fillna(0.0).astype(float)
    
    return df


# -----------------------------------
# FILL EXCEL TEMPLATE AUTOMATICALLY
# -----------------------------------
def fill_excel_template(template_file, patient, member, provider, scan_row):

    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    # --------- DETECT FIELDS AUTOMATICALLY ---------

    patient_cell = None
    member_cell = None
    provider_cell = None
    scan_start_row = None
    desc_col = tarif_col = modi_col = qty_col = amt_col = None
    total_cell = None

    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                val = str(cell.value).strip().upper()

                if "FOR PATIENT" in val:
                    patient_cell = ws.cell(row=cell.row, column=cell.column + 1)

                if "MEMBER NUMBER" in val:
                    member_cell = ws.cell(row=cell.row, column=cell.column + 1)

                if "MEDICAL EXAMINATION" in val:
                    provider_cell = ws.cell(row=cell.row, column=cell.column + 1)

                if val == "DESCRIPTION":
                    scan_start_row = cell.row + 1
                    desc_col = cell.column
                    tarif_col = cell.column + 1
                    modi_col = cell.column + 2
                    qty_col = cell.column + 3
                    amt_col = cell.column + 4

                if val == "TOTAL":
                    # The total amount cell in your screenshot is approx. 6 columns to the right of "TOTAL"
                    total_cell = ws.cell(row=cell.row, column=cell.column + 6)

    # Validate detection
    if not all([patient_cell, member_cell, provider_cell, scan_start_row, desc_col, tarif_col, modi_col, qty_col, amt_col, total_cell]):
        st.error("Could not detect all required fields in the quotation template. Please check the template format.")
        return None

    # --------- WRITE PATIENT INFO ---------
    patient_cell.value = patient
    member_cell.value = member
    provider_cell.value = provider

    # --------- WRITE SCAN DATA INTO TABLE ---------
    ws.cell(row=scan_start_row, column=desc_col, value=scan_row["EXAMINATION"])
    ws.cell(row=scan_start_row, column=tarif_col, value=scan_row["TARIFF"])
    ws.cell(row=scan_start_row, column=modi_col, value=scan_row["MODIFIER"])
    ws.cell(row=scan_start_row, column=qty_col, value=int(scan_row["QUANTITY"]))
    ws.cell(row=scan_start_row, column=amt_col, value=float(scan_row["AMOUNT"]))

    # --------- WRITE TOTAL ---------
    total_cell.value = float(scan_row["AMOUNT"])

    # --------- OUTPUT ---------
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# -----------------------------------
# STREAMLIT USER INTERFACE
# -----------------------------------

st.title("📄 Medical Quotation Generator (Excel Template Auto-Detection)")

charge_file = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
template_file = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

if charge_file and template_file:
    if st.button("Continue"):
        df = load_charge_sheet(charge_file)
        if df is not None:
            st.success("Charge sheet loaded!")

            st.subheader("Patient Information")
            col1, col2 = st.columns(2)

            with col1:
                patient = st.text_input("Patient Name")
                member = st.text_input("Medical Aid Number")

            with col2:
                provider = st.text_input("Medical Aid Provider", value="CIMAS")

            st.subheader("Select Scan")
            scan_list = df["EXAMINATION"].unique().tolist()
            selected_scan = st.selectbox("Choose Scan", scan_list)

            if selected_scan:
                scan_row = df[df["EXAMINATION"] == selected_scan].iloc[0]

                st.write("### Scan Details:")
                st.write(f"**Tariff:** {scan_row['TARIFF']}")
                st.write(f"**Modifier:** {scan_row['MODIFIER']}")
                st.write(f"**Quantity:** {scan_row['QUANTITY']}")
                st.write(f"**Amount:** {scan_row['AMOUNT']}")

                if st.button("Generate Quotation"):
                    output = fill_excel_template(
                        template_file,
                        patient,
                        member,
                        provider,
                        scan_row
                    )
                    if output:
                        st.success("Quotation generated!")

                        st.download_button(
                            "Download Excel Quotation",
                            data=output,
                            file_name=f"quotation_{patient}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
else:
    st.info("Upload your charge sheet and quotation template to continue.")
