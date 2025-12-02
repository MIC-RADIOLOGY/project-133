import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import difflib


# ---------------------------------------------------------
# FIND A CELL BY KEYWORD IN THE TEMPLATE
# ---------------------------------------------------------
def find_cell(sheet, keyword):
    keyword = keyword.lower()
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and keyword in str(cell.value).lower():
                return cell.row, cell.column
    return None


# ---------------------------------------------------------
# GET TARIFF USING YOUR EXACT HEADER STRUCTURE
# ---------------------------------------------------------
def get_tariff(scan_type, charge_file):
    charges = pd.read_excel(charge_file, sheet_name=None)

    for sheet_name, df in charges.items():

        # Must contain EXAMINATION column
        exam_col = None
        for col in df.columns:
            if "examination" in col.lower():
                exam_col = col
                break

        if not exam_col:
            continue

        # Price is always the last column (CIMAS USD)
        price_col = df.columns[-1]

        # Also keep TARIFF and QTY if available
        tariff_col = None
        qty_col = None
        for col in df.columns:
            if "tariff" in col.lower():
                tariff_col = col
            if "qty" in col.lower():
                qty_col = col

        # LOWER CASE COLUMN to match search
        df["__lower_exam"] = df[exam_col].astype(str).str.lower()

        user_lower = scan_type.lower()

        # Fuzzy match the scan type to EXAMINATION column
        matches = difflib.get_close_matches(user_lower, df["__lower_exam"], n=1, cutoff=0.4)

        if matches:
            row = df[df["__lower_exam"] == matches[0]].iloc[0]

            return {
                "EXAMINATION": row[exam_col],
                "TARIFF": row[tariff_col] if tariff_col else "",
                "QTY": row[qty_col] if qty_col else 1,
                "AMOUNT": row[price_col],   # ALWAYS LAST COLUMN
            }

    return None


# ---------------------------------------------------------
# FILL TEMPLATE EXACTLY AS IS
# ---------------------------------------------------------
def fill_template(template_file, patient, medaid, scan, tariff_data):
    wb = load_workbook(template_file)
    sheet = wb.active

    # Locate auto-detected fields
    patient_cell = find_cell(sheet, "patient")
    medaid_cell  = find_cell(sheet, "medical")
    scan_cell    = find_cell(sheet, "scan")
    tariff_cell  = find_cell(sheet, "examination")

    if not (patient_cell and medaid_cell and scan_cell and tariff_cell):
        st.error("Template missing required headings.")
        return None

    # Fill patient details
    sheet.cell(patient_cell[0], patient_cell[1] + 1).value = patient
    sheet.cell(medaid_cell[0], medaid_cell[1] + 1).value  = str(medaid)
    sheet.cell(scan_cell[0], scan_cell[1] + 1).value      = scan

    # Fill tariff table
    start_row = tariff_cell[0] + 1
    col = tariff_cell[1]

    sheet.cell(start_row, col).value       = tariff_data["EXAMINATION"]
    sheet.cell(start_row, col + 1).value   = tariff_data["TARIFF"]
    sheet.cell(start_row, col + 2).value   = tariff_data["QTY"]
    sheet.cell(start_row, col + 3).value   = tariff_data["AMOUNT"]

    # Save to memory
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ---------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------
st.title("AI Radiology Quotation Generator (Excel)")

template_file = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])
charge_file   = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
medaid  = st.text_input("Medical Aid Number")
scan    = st.text_input("Scan Type (as written under EXAMINATION)")

if st.button("Generate Quotation"):
    if not template_file or not charge_file:
        st.error("Upload both the quotation template AND the charge sheet.")
    elif not (patient and medaid and scan):
        st.error("Complete all fields.")
    else:
        tariff_data = get_tariff(scan, charge_file)

        if not tariff_data:
            st.error("❌ Scan not found. Try typing part of the EXAMINATION name.")
        else:
            output = fill_template(template_file, patient, medaid, scan, tariff_data)
            if output:
                st.success("Quotation Ready!")
                st.download_button(
                    "Download Final Excel Quotation",
                    output,
                    file_name="quotation_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
