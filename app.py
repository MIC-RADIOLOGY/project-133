import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import difflib


# ---------------------------------------------------------
# FIND A CELL BY KEYWORD
# ---------------------------------------------------------
def find_cell(sheet, keyword):
    keyword = keyword.lower()
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and keyword in str(cell.value).lower():
                return cell.row, cell.column
    return None


# ---------------------------------------------------------
# GET TARIFF WITH FLEXIBLE COLUMN SEARCH + FUZZY MATCH
# ---------------------------------------------------------
def get_tariff(scan_type, charge_file):
    charges = pd.read_excel(charge_file, sheet_name=None)

    for sheet_name, df in charges.items():

        # FIX: detect the scan column automatically
        possible_columns = ["scan", "scan_name", "service", "description", "procedure", "exam", "item"]

        scan_col = None
        for col in df.columns:
            if any(p in col.lower() for p in possible_columns):
                scan_col = col
                break

        if not scan_col:
            continue  # sheet has no valid scan column → skip

        # FIX: fuzzy match instead of exact match
        df["__lower"] = df[scan_col].astype(str).str.lower()
        user_lower = scan_type.lower()

        # find closest match
        matches = difflib.get_close_matches(user_lower, df["__lower"], n=1, cutoff=0.4)
        if matches:
            match = df[df["__lower"] == matches[0]].iloc[0]
            return match.to_dict()

    return None



# ---------------------------------------------------------
# FILL TEMPLATE EXACTLY AS-IS
# ---------------------------------------------------------
def fill_template(template_file, patient, medaid, scan, tariff_data):
    wb = load_workbook(template_file)
    sheet = wb.active

    patient_cell = find_cell(sheet, "patient")
    medaid_cell = find_cell(sheet, "medical")
    scan_cell = find_cell(sheet, "scan")
    tariff_cell = find_cell(sheet, "tariff")

    if not (patient_cell and medaid_cell and scan_cell and tariff_cell):
        st.error("Template missing required headings.")
        return None

    # --- FIX: Always write medical aid number as TEXT ---
    medaid_text = str(medaid)

    # Fill fields
    sheet.cell(patient_cell[0], patient_cell[1] + 1).value = patient
    sheet.cell(medaid_cell[0], medaid_cell[1] + 1).value = medaid_text
    sheet.cell(scan_cell[0], scan_cell[1] + 1).value = scan

    # Tariff table
    start_row = tariff_cell[0] + 1
    col_tariff = tariff_cell[1]

    for i, (key, value) in enumerate(tariff_data.items()):
        if key != "__lower":  # skip helper column
            sheet.cell(start_row + i, col_tariff).value = key
            sheet.cell(start_row + i, col_tariff + 1).value = value

    # Save output
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output



# ---------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------
st.title("AI Radiology Quotation Generator (Excel Version)")

template_file = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])
charge_file = st.fi_
