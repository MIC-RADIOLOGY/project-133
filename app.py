import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# -----------------------------------
# LOAD CHARGE SHEET
# -----------------------------------
def load_charge_sheet(file):
    df = pd.read_excel(file)
    df.columns = [c.strip().upper() for c in df.columns]
    df = df[df["TARRIF"].notna()]
    df["EXAMINATION"] = df["EXAMINATION"].astype(str)
    df["TARRIF"] = df["TARRIF"].astype(str)
    df["MODIFIER"] = df["MODIFIER"].astype(str)
    return df


# -----------------------------------
# FILL EXCEL TEMPLATE AUTOMATICALLY
# -----------------------------------
def fill_excel_template(template_file, patient, member, provider, scan_row):

    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    # --------- DETECT FIELDS AUTOMATICALLY ---------

    # Locate patient name cell
    patient_cell = None
    member_cell = None

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

                # Scan table header
                if val == "DESCRIPTION":
                    scan_start_row = cell.row + 1
                    desc_col = cell.column
                    tarif_col = cell.column + 1
                    modi_col = cell.column + 2
                    qty_col = cell.column + 3
                    amt_col = cell.column + 4

                # Total detection
                if val == "TOTAL":
                    total_cell = ws.cell(row=cell.row, column=cell.column + 6)

    # --------- WRITE PATIENT INFO ---------
    if patient_cell:
        patient_cell.value = patient

    if member_cell:
        member_cell.value = member

    if provider_cell:
        provider_cell.value = provider

    # --------- WRITE SCAN DATA INTO TABLE ---------
    ws.cell(row=scan_start_row, column=desc_col, value=scan_row["EXAMINATION"])
    ws.cell(row=scan_start_row, column=tarif_col, value=scan_row["TARRIF"])
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

    df = load_charge_sheet(charge_file)
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

            st.success("Quotation generated!")

            st.download_button(
                "Download Excel Quotation",
                data=output,
                file_name=f"quotation_{patient}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("Upload your charge sheet and template to continue.")
