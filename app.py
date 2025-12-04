import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# -----------------------------------
# LOAD CHARGE SHEET (no header in Excel)
# -----------------------------------
def load_charge_sheet(file):
    df = pd.read_excel(file, header=None)
    df.columns = ["EXAMINATION", "TARRIF", "MODIFIER", "QUANTITY", "AMOUNT"]
    st.write("DEBUG: Columns after assigning:", df.columns.tolist())
    df = df[df["TARRIF"].notna()]
    df["EXAMINATION"] = df["EXAMINATION"].astype(str)
    df["TARRIF"] = df["TARRIF"].astype(str)
    df["MODIFIER"] = df["MODIFIER"].astype(str)
    df["QUANTITY"] = pd.to_numeric(df["QUANTITY"], errors='coerce').fillna(0).astype(int)
    df["AMOUNT"] = pd.to_numeric(df["AMOUNT"], errors='coerce').fillna(0.0).astype(float)
    return df

# -----------------------------------
# FILL EXCEL TEMPLATE AUTOMATICALLY
# -----------------------------------
def fill_excel_template(template_file, patient, member, provider, scan_row):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    # Detect fields
    patient_cell = member_cell = provider_cell = None
    scan_start_row = desc_col = tarif_col = modi_col = qty_col = amt_col = total_cell = None

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
                    total_cell = ws.cell(row=cell.row, column=cell.column + 6)

    if not all([patient_cell, member_cell, provider_cell, scan_start_row, desc_col, tarif_col, modi_col, qty_col, amt_col, total_cell]):
        st.error("Could not detect all required fields in the quotation template. Please check the template format.")
        return None

    # Write patient info
    patient_cell.value = patient
    member_cell.value = member
    provider_cell.value = provider

    # Write scan data
    ws.cell(row=scan_start_row, column=desc_col, value=scan_row["EXAMINATION"])
    ws.cell(row=scan_start_row, column=tarif_col, value=scan_row["TARRIF"])
    ws.cell(row=scan_start_row, column=modi_col, value=scan_row["MODIFIER"])
    ws.cell(row=scan_start_row, column=qty_col, value=int(scan_row["QUANTITY"]))
    ws.cell(row=scan_start_row, column=amt_col, value=float(scan_row["AMOUNT"]))

    # Write total
    total_cell.value = float(scan_row["AMOUNT"])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -----------------------------------
# STREAMLIT INTERFACE
# -----------------------------------

st.title("📄 Medical Quotation Generator")

# -----------------------------
# Persistent file uploads
# -----------------------------
if "charge_file" not in st.session_state:
    st.session_state.charge_file = None
if "template_file" not in st.session_state:
    st.session_state.template_file = None

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
if uploaded_charge is not None:
    st.session_state.charge_file = uploaded_charge

uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])
if uploaded_template is not None:
    st.session_state.template_file = uploaded_template

charge_file = st.session_state.charge_file
template_file = st.session_state.template_file

# -----------------------------
# Persistent patient info
# -----------------------------
if "patient_input" not in st.session_state:
    st.session_state.patient_input = ""
if "member_input" not in st.session_state:
    st.session_state.member_input = ""
if "provider_input" not in st.session_state:
    st.session_state.provider_input = "CIMAS"

# Only show Continue button if both files uploaded
if charge_file and template_file:
    if st.button("Continue"):
        df = load_charge_sheet(charge_file)
        if df is not None:
            st.success("Charge sheet loaded!")

            st.subheader("Patient Information")
            col1, col2 = st.columns(2)

            with col1:
                st.text_input("Patient Name", key="patient_input")
                st.text_input("Medical Aid Number", key="member_input")
            with col2:
                st.text_input("Medical Aid Provider", key="provider_input", value=st.session_state.provider_input)

            patient = st.session_state.patient_input
            member = st.session_state.member_input
            provider = st.session_state.provider_input

            st.subheader("Select Scan")
            scan_list = df["EXAMINATION"].unique().tolist()
            selected_scan = st.selectbox("Choose Scan", scan_list, key="scan_select")

            if selected_scan:
                scan_row = df[df["EXAMINATION"] == selected_scan].iloc[0]

                st.write("### Scan Details:")
                st.write(f"**Tariff:** {scan_row['TARRIF']}")
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
