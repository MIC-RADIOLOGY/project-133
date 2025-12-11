import streamlit as st
import pandas as pd
import openpyxl
import io
import math
from typing import Optional

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config / heuristics ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN", "MRI", "X-RAY",
    "ECG", "MAMMOGRAPHY", "DEXA", "PANOREX", "OTHERS"
}

# ---------- Utility ----------
def normalize(text: Optional[str]) -> str:
    if text is None:
        return ""
    return str(text).strip().upper()

# ---------- Parse charge sheet ----------
def parse_charge_sheet(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [normalize(c) for c in df.columns]

    expected_cols = {"DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"}
    missing_cols = expected_cols - set(df.columns)
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")

    # We only keep valid lines
    cleaned = df[[
        "DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"
    ]].copy()

    # Normalise values
    cleaned["DESCRIPTION"] = cleaned["DESCRIPTION"].astype(str).str.strip()
    cleaned["TARIFF"] = cleaned["TARIFF"].astype(str).str.strip()
    cleaned["MODIFIER"] = cleaned["MODIFIER"].astype(str).str.strip()

    # Convert QTY + FEES + AMOUNT to numbers
    cleaned["QTY"] = pd.to_numeric(cleaned["QTY"], errors="coerce").fillna(1)
    cleaned["FEES"] = pd.to_numeric(cleaned["FEES"], errors="coerce").fillna(0)
    cleaned["AMOUNT"] = pd.to_numeric(cleaned["AMOUNT"], errors="coerce").fillna(
        cleaned["FEES"] * cleaned["QTY"]
    )

    # Remove empty lines
    cleaned = cleaned[cleaned["DESCRIPTION"] != ""]
    cleaned = cleaned[cleaned["TARIFF"] != ""]

    return cleaned.reset_index(drop=True)

# ---------- Fill Excel Template ----------
def fill_excel_template(template_file, patient, member, provider, items_df):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    # Replace header fields (REMOVE OLD NAME AND INSERT NEW ONE)
    ws["A2"] = f"ATT: {provider}"
    ws["A3"] = f"FOR PATIENT: {patient}"
    ws["A4"] = f"MEMBER NUMBER: {member}"

    # Starting row for items
    start_row = 21
    current_row = start_row

    # Clear previous lines but DO NOT remove the blue total line
    for r in range(start_row, 60):
        for c in range(1, 8):
            if ws.cell(r, c).fill.start_color.index not in ("00000000", None):
                # skip styled rows (blue)
                continue
            ws.cell(r, c).value = None

    # Write new items
    for _, row in items_df.iterrows():
        ws.cell(current_row, 1).value = row["DESCRIPTION"]
        ws.cell(current_row, 2).value = row["TARIFF"]
        ws.cell(current_row, 3).value = row["MODIFIER"]
        ws.cell(current_row, 4).value = row["QTY"]
        ws.cell(current_row, 5).value = row["FEES"]
        ws.cell(current_row, 6).value = row["AMOUNT"]
        current_row += 1

    # Total must remain in G22
    ws["G22"] = f"=SUM(F{start_row}:F{current_row - 1})"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ---------- Streamlit UI ----------
st.title("Medical Quotation Generator")

uploaded_template = st.file_uploader("Upload Excel Template", type=["xlsx"])
uploaded_charge = st.file_uploader("Upload Charge Sheet", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Provider (ATT:)")

if uploaded_charge:
    try:
        df_raw = pd.read_excel(uploaded_charge)
        parsed = parse_charge_sheet(df_raw)
        st.success("Charge sheet parsed successfully.")
        st.dataframe(parsed)

        if uploaded_template:
            if st.button("Generate Quotation"):
                output_excel = fill_excel_template(
                    uploaded_template,
                    patient,
                    member,
                    provider,
                    parsed
                )

                st.download_button(
                    label="Download Quotation",
                    data=output_excel,
                    file_name="quotation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error parsing charge sheet: {e}")

else:
    st.info("Upload a charge sheet to begin.")
