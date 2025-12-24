import streamlit as st
import pandas as pd
import openpyxl
import io
from copy import copy
from datetime import datetime

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ------------------------------------------------------------
# CONSTANT COLUMN MAP (MERGE-SAFE)
# ------------------------------------------------------------
COL_MAP = {
    "DESCRIPTION": "A",
    "TARIFF": "B",
    "MODIFIER": "C",
    "QTY": "E",
    "FEES": "F",
    "AMOUNT": "G",
}

START_ROW = 22
TOTAL_ROW = 22  # total must appear ONLY here

# ------------------------------------------------------------
# FILE UPLOAD
# ------------------------------------------------------------
uploaded_file = st.file_uploader("Upload Excel Tariff File", type=["xlsx"])

if not uploaded_file:
    st.stop()

df = pd.read_excel(uploaded_file)
df.columns = [c.strip().upper() for c in df.columns]

required = {"DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"}
missing = required - set(df.columns)

if missing:
    st.error(f"Missing columns: {missing}")
    st.stop()

# ------------------------------------------------------------
# REVIEW TABLE
# ------------------------------------------------------------
st.subheader("Quotation Review")
st.dataframe(df[["DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"]])

# ------------------------------------------------------------
# LOAD TEMPLATE
# ------------------------------------------------------------
template_path = "quotation_template.xlsx"  # your template
wb = openpyxl.load_workbook(template_path)
ws = wb.active

# ------------------------------------------------------------
# CLEAR PREVIOUS ROWS
# ------------------------------------------------------------
for r in range(START_ROW, START_ROW + 30):
    for col in COL_MAP.values():
        ws[f"{col}{r}"].value = None

# ------------------------------------------------------------
# WRITE ROWS (MODIFIER FIXED)
# ------------------------------------------------------------
row_ptr = START_ROW
total_amount = 0

for _, row in df.iterrows():
    ws[f"{COL_MAP['DESCRIPTION']}{row_ptr}"] = row["DESCRIPTION"]
    ws[f"{COL_MAP['TARIFF']}{row_ptr}"] = row["TARIFF"]
    ws[f"{COL_MAP['MODIFIER']}{row_ptr}"] = row["MODIFIER"]  # FIXED
    ws[f"{COL_MAP['QTY']}{row_ptr}"] = row["QTY"]
    ws[f"{COL_MAP['FEES']}{row_ptr}"] = row["FEES"]
    ws[f"{COL_MAP['AMOUNT']}{row_ptr}"] = row["AMOUNT"]

    total_amount += float(row["AMOUNT"])
    row_ptr += 1

# ------------------------------------------------------------
# REMOVE SUBTOTAL / WRITE TOTAL ONLY
# ------------------------------------------------------------
ws[f"{COL_MAP['AMOUNT']}{TOTAL_ROW}"] = round(total_amount, 2)

# ------------------------------------------------------------
# DATE (NUMERIC FORMAT)
# ------------------------------------------------------------
ws["G8"] = datetime.today().strftime("%d/%m/%Y")

# ------------------------------------------------------------
# SAVE FILE
# ------------------------------------------------------------
output = io.BytesIO()
wb.save(output)
output.seek(0)

st.download_button(
    "Download Quotation",
    data=output,
    file_name="quotation_generated.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
