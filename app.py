import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import difflib

def find_cell_any(sheet, keywords):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value:
                text = str(cell.value).lower()
                for word in keywords:
                    if word.lower() in text:
                        for merged in sheet.merged_cells.ranges:
                            if cell.coordinate in merged:
                                return merged.min_row, merged.min_col
                        return cell.row, cell.column
    return None

def safe_write(sheet, row, col, value):
    for merged in sheet.merged_cells.ranges:
        if sheet.cell(row, col).coordinate in merged:
            row, col = merged.min_row, merged.min_col
            break
    sheet.cell(row, col, value)

def get_tariffs_for_scan(scan_type, charge_file):
    charges = pd.read_excel(charge_file, sheet_name=None)
    results = []

    for sheet_name, df in charges.items():
        exam_col = None
        for col in df.columns:
            if "examination" in col.lower():
                exam_col = col
                break
        if not exam_col:
            continue

        df["__lower_exam"] = df[exam_col].astype(str).str.lower()
        user_lower = scan_type.lower()

        # Filter rows where examination contains the scan_type text (case-insensitive)
        matched_rows = df[df["__lower_exam"].str.contains(user_lower, na=False)]

        if not matched_rows.empty:
            # Add necessary columns if missing
            for col in ["TARIFF", "MOD", "QTY", "FEES", "AMOUNT"]:
                if col not in matched_rows.columns:
                    matched_rows[col] = ""

            for _, row in matched_rows.iterrows():
                results.append(row.to_dict())
            return results
    return None

def fill_template_multiple(template_file, patient, medaid, scan, tariffs):
    wb = load_workbook(template_file)
    sheet = wb.active

    # Find patient, medaid, scan cells
    patient_cell = find_cell_any(sheet, ["patient", "name", "client"])
    medaid_cell  = find_cell_any(sheet, ["medical", "member", "aid", "scheme"])
    scan_cell    = find_cell_any(sheet, ["scan", "exam", "procedure", "service", "investigation"])
    
    header_cell = find_cell_any(sheet, ["description"])
    if not header_cell:
        st.error("Template missing DESCRIPTION header.")
        return None

    header_row = header_cell[0]

    # Map headers to columns in the header row
    headers = {}
    for cell in sheet[header_row]:
        if cell.value:
            val = str(cell.value).strip().lower()
            headers[val] = cell.column

    required_cols = ["description", "tarrif", "mod", "qty", "fees", "amount"]
    for col in required_cols:
        if col not in headers:
            st.error(f"Template missing column header: {col.upper()}")
            return None

    # Fill patient info
    safe_write(sheet, patient_cell[0], patient_cell[1] + 1, patient)
    safe_write(sheet, medaid_cell[0], medaid_cell[1] + 1, str(medaid))
    safe_write(sheet, scan_cell[0], scan_cell[1] + 1, scan)

    start_row = header_row + 1

    for i, tariff_row in enumerate(tariffs):
        row_idx = start_row + i

        safe_write(sheet, row_idx, headers["description"], str(tariff_row.get("EXAMINATION", "")))
        safe_write(sheet, row_idx, headers["tarrif"], tariff_row.get("TARIFF", ""))
        safe_write(sheet, row_idx, headers["mod"], tariff_row.get("MOD", ""))
        safe_write(sheet, row_idx, headers["qty"], tariff_row.get("QTY", 1))
        safe_write(sheet, row_idx, headers["fees"], tariff_row.get("FEES", ""))
        safe_write(sheet, row_idx, headers["amount"], tariff_row.get("AMOUNT", ""))

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

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
        tariffs = get_tariffs_for_scan(scan, charge_file)
        if not tariffs:
            st.error("❌ Scan not found. Try typing part of the EXAMINATION name.")
        else:
            output = fill_template_multiple(template_file, patient, medaid, scan, tariffs)
            if output:
                st.success("Quotation Ready!")
                st.download_button(
                    "Download Final Excel Quotation",
                    output,
                    file_name="quotation_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
