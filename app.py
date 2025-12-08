import streamlit as st
import pandas as pd
import openpyxl
import io
import math
from typing import Optional

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# -----------------------------------------------------------
# CATEGORY DEFINITIONS (UPDATED WITH MRI)
# -----------------------------------------------------------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN",
    "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND",
    "MRI"
}

GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# -----------------------------------------------------------
# Load charge sheet + parse into structured dictionary
# -----------------------------------------------------------
def load_charge_sheet(file) -> dict:
    xls = pd.ExcelFile(file)
    services = {}

    for sheet in xls.sheet_names:
        df = pd.read_excel(file, sheet_name=sheet, header=None)

        current_category = None

        for i in range(len(df)):
            raw_value = str(df.iloc[i, 0]).strip()

            if raw_value.upper() in MAIN_CATEGORIES:
                current_category = raw_value.upper()
                if current_category not in services:
                    services[current_category] = []
                continue

            if current_category and raw_value != "" and raw_value.upper() not in GARBAGE_KEYS:

                description = raw_value
                usd_col = None

                for col in df.columns:
                    cell = df.iloc[i, col]
                    if isinstance(cell, (int, float)) and not math.isnan(cell):
                        usd_col = float(cell)
                        break

                if usd_col is None:
                    continue

                services[current_category].append({
                    "description": description,
                    "amount": usd_col
                })

    return services

# -----------------------------------------------------------
# Match tariff from user input
# -----------------------------------------------------------
def find_tariff(services: dict, search_text: str) -> Optional[dict]:
    if not search_text:
        return None
    search_text = search_text.lower()

    for cat, items in services.items():
        for item in items:
            if search_text in item["description"].lower():
                return {
                    "category": cat,
                    "description": item["description"],
                    "amount": item["amount"]
                }
    return None

# -----------------------------------------------------------
# Fill quotation template
# -----------------------------------------------------------
def fill_template(template_file, patient_name, member_number, items):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    ws["C2"] = patient_name
    ws["C3"] = member_number

    start_row = 6
    row = start_row

    for tariff in items:
        ws.cell(row=row, column=1).value = tariff["description"]
        ws.cell(row=row, column=7).value = tariff["amount"]
        row += 1

    total_cell = "G22"
    ws[total_cell].value = f"=SUM(G{start_row}:G{row-1})"

    return wb

# -----------------------------------------------------------
# STREAMLIT UI
# -----------------------------------------------------------
st.title("Medical Quotation Generator – Final Version (MRI Fixed)")

st.subheader("1. Upload Charge Sheet")
charge_file = st.file_uploader("Upload the charge sheet (Excel)", type=["xlsx"])

st.subheader("2. Upload Quotation Template")
template_file = st.file_uploader("Upload the quotation template (Excel)", type=["xlsx"])

if charge_file and template_file:
    services = load_charge_sheet(charge_file)
    st.success("Charge sheet processed successfully.")

    st.subheader("3. Patient Details")
    patient_name = st.text_input("Patient Name")
    member_number = st.text_input("Member Number")

    st.subheader("4. Enter Tariffs (one per line)")
    raw_tariffs = st.text_area("Enter tariffs", height=200)

    if st.button("Generate Quotation"):
        lines = [x.strip() for x in raw_tariffs.split("\n") if x.strip()]
        matched_items = []

        for line in lines:
            match = find_tariff(services, line)
            if match:
                matched_items.append(match)
            else:
                st.warning(f"No matching tariff found: {line}")

        if matched_items:
            wb = fill_template(template_file, patient_name, member_number, matched_items)
            output = io.BytesIO()
            wb.save(output)
            st.download_button(
                label="Download Final Quotation",
                data=output.getvalue(),
                file_name="Quotation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No valid tariffs found. Please check inputs.")
