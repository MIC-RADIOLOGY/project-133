import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# -----------------------------------------------------
# 1. FIND CELL BY KEYWORD (SEARCH THE TEMPLATE)
# -----------------------------------------------------
def find_cell(sheet, keyword):
    keyword = keyword.lower()
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and keyword in str(cell.value).lower():
                return cell.row, cell.column
    return None

# -----------------------------------------------------
# 2. GET TARIFF FROM MULTIPLE SHEETS
# -----------------------------------------------------
def get_tariff(scan_type):
    charges = pd.read_excel("charges.xlsx", sheet_name=None)

    for sheet_name, df in charges.items():
        if "Scan_Name" not in df.columns:
            continue
        
        match = df[df["Scan_Name"].str.lower() == scan_type.lower()]
        if not match.empty:
            return match.to_dict(orient="records")[0]
    
    return None  # not found

# -----------------------------------------------------
# 3. FILL TEMPLATE EXACTLY AS-IS
# -----------------------------------------------------
def fill_quotation_template(patient, medaid, scan, tariff_data):
    wb = load_workbook("quotation_template.xlsx")
    sheet = wb.active

    # Locate fields automatically
    patient_cell = find_cell(sheet, "patient")
    medaid_cell = find_cell(sheet, "medical")
    scan_cell = find_cell(sheet, "scan")
    tariff_header = find_cell(sheet, "tariff")  # header row

    if not (patient_cell and medaid_cell and scan_cell and tariff_header):
        st.error("Template structure missing required keywords.")
        return None

    # Fill top fields
    sheet.cell(patient_cell[0], patient_cell[1] + 1).value = patient
    sheet.cell(medaid_cell[0], medaid_cell[1] + 1).value = medaid
    sheet.cell(scan_cell[0], scan_cell[1] + 1).value = scan

    # Fill tariff table (starting below header)
    start_row = tariff_header[0] + 1

    col_tariff = tariff_header[1]
    col_price = col_tariff + 1

    for i, (key, value) in enumerate(tariff_data.items()):
        sheet.cell(start_row + i, col_tariff).value = key
        sheet.cell(start_row + i, col_price).value = value

    # Save output file
    output_file = "quotation_output.xlsx"
    wb.save(output_file)
    return output_file

# -----------------------------------------------------
# 4. STREAMLIT USER INTERFACE
# -----------------------------------------------------
st.title("Radiology Quotation Generator (Excel)")

patient = st.text_input("Patient Name")
medaid = st.text_input("Medical Aid Number")
scan = st.text_input("Type of Scan (must match charges sheet)")

if st.button("Generate Quotation"):
    if not (patient and medaid and scan):
        st.error("Please fill all fields.")
    else:
        tariff_data = get_tariff(scan)
        if not tariff_data:
            st.error("Scan type not found in charge sheet.")
        else:
            output = fill_quotation_template(patient, medaid, scan, tariff_data)
            if output:
                with open(output, "rb") as f:
                    st.download_button(
                        "Download Completed Quotation",
                        f,
                        file_name="quotation_output.xlsx"
                    )
