import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import difflib

# ---------------------------
# Flexible cell finder
# ---------------------------
def find_cell_any(sheet, keywords):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value:
                text = str(cell.value).lower()
                for word in keywords:
                    if word.lower() in text:
                        return cell.row, cell.column
    return None

# ---------------------------
# Get tariff from charge sheet
# ---------------------------
def get_tariff(scan_type, charge_file):
    charges = pd.read_excel(charge_file, sheet_name=None)

    for sheet_name, df in charges.items():
        # EXAMINATION column
        exam_col = None
        for col in df.columns:
            if "examination" in col.lower():
                exam_col = col
                break
        if not exam_col:
            continue

        price_col = df.columns[-1]  # last column is always amount
        tariff_col = None
        qty_col = None
        for col in df.columns:
            if "tariff" in col.lower():
                tariff_col = col
            if "qty" in col.lower():
                qty_col = col

        df["__lower_exam"] = df[exam_col].astype(str).str.lower()
        user_lower = scan_type.lower()

        matches = difflib.get_close_matches(user_lower, df["__lower_exam"], n=1, cutoff=0.4)
        if matches:
            row = df[df["__lower_exam"] == matches[0]].iloc[0]
            return {
                "EXAMINATION": row[exam_col],
                "TARIFF": row[tariff_col] if tariff_col else "",
                "QTY": row[qty_col] if qty_col else 1,
                "AMOUNT": row[price_col],
            }
    return None

# ---------------------------
# Fill template
# ---------------------------
def fill_template(template_file, patient, medaid, scan, tariff_data):
    wb = load_workbook(template_file)
    sheet = wb.active

    # Flexible keywords
    patient_cell = find_cell_any(sheet, ["patient", "name", "client"])
    medaid_cell  = find_cell_any(sheet, ["medical", "member", "aid", "scheme"])
    scan_cell    = find_cell_any(sheet, ["scan", "exam", "procedure", "service", "investigation"])
    tariff_cell  = find_cell_any(sheet, ["examination", "service", "procedure", "item", "description"])

    # Debug output to see what was found
    st.write("Found patient cell:", patient_cell)
    st.write("Found medical aid cell:", medaid_cell)
    st.write("Found scan cell:", scan_cell)
    st.write("Found tariff cell:", tariff_cell)

    if not (patient_cell and medaid_cell and scan_cell and tariff_cell):
        st.error("Template missing required headings. Make sure it contains patient, medical, scan, and examination/service headers.")
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

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ---------------------------
# Streamlit UI
# ---------------------------
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
