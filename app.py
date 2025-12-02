import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

# ---------------------------------------------------------
# FIND A CELL BY KEYWORD (EVEN IF THE CELL IS MERGED)
# ---------------------------------------------------------
def find_cell(sheet, keyword):
    keyword = keyword.lower()
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and keyword in str(cell.value).lower():
                return cell.row, cell.column
    return None


# ---------------------------------------------------------
# GET TARIFF FROM MULTIPLE SHEETS
# ---------------------------------------------------------
def get_tariff(scan_type, charge_file):
    charges = pd.read_excel(charge_file, sheet_name=None)

    for sheet_name, df in charges.items():
        if "Scan_Name" not in df.columns:
            continue

        match = df[df["Scan_Name"].str.lower() == scan_type.lower()]
        if not match.empty:
            return match.to_dict(orient="records")[0]

    return None


# ---------------------------------------------------------
# FILL THE QUOTATION TEMPLATE EXACTLY AS-IS
# ---------------------------------------------------------
def fill_template(template_file, patient, medaid, scan, tariff_data):
    # Load workbook from uploaded file
    wb = load_workbook(template_file)
    sheet = wb.active

    # Find key fields
    patient_cell = find_cell(sheet, "patient")
    medaid_cell = find_cell(sheet, "medical")
    scan_cell = find_cell(sheet, "scan")
    tariff_cell = find_cell(sheet, "tariff")

    if not (patient_cell and medaid_cell and scan_cell and tariff_cell):
        st.error("Template missing required headings.")
        return None

    # Fill data next to detected fields
    sheet.cell(patient_cell[0], patient_cell[1] + 1).value = patient

    # MEDICAL AID AS NUMBER OR TEXT
    try:
        medaid_num = int(medaid)
        sheet.cell(medaid_cell[0], medaid_cell[1] + 1).value = medaid_num
    except:
        sheet.cell(medaid_cell[0], medaid_cell[1] + 1).value = medaid

    sheet.cell(scan_cell[0], scan_cell[1] + 1).value = scan

    # Insert tariff table
    start_row = tariff_cell[0] + 1
    tariff_col = tariff_cell[1]

    for i, (key, value) in enumerate(tariff_data.items()):
        sheet.cell(start_row + i, tariff_col).value = key
        sheet.cell(start_row + i, tariff_col + 1).value = value

    # Save to memory
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ---------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------
st.title("AI Radiology Quotation Generator (Excel Auto-Template)")

template_file = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])
charge_file = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
medaid = st.text_input("Medical Aid Number")
scan = st.text_input("Scan Type (exact name)")

if st.button("Generate Quotation"):
    if not template_file or not charge_file:
        st.error("Upload both the quotation template and charge sheet.")
    elif not (patient and medaid and scan):
        st.error("Fill all patient fields.")
    else:
        tariff_data = get_tariff(scan, charge_file)
        if not tariff_data:
            st.error("Scan type not found in charge sheet.")
        else:
            output = fill_template(template_file, patient, medaid, scan, tariff_data)
            if output:
                st.download_button(
                    "Download Finished Quotation",
                    output,
                    file_name="quotation_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
