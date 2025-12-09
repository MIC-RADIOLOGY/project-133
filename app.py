import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# -------------------------------------------------------
# Upload Inputs
# -------------------------------------------------------
st.header("Medical Quotation Generator")

uploaded_charge_sheet = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Provider Name")
quote_date = st.date_input("Quotation Date")

st.write("Select items that will appear in the quotation:")

if uploaded_charge_sheet:
    df = pd.read_excel(uploaded_charge_sheet)

    # Ensure expected columns
    expected_columns = {"Description", "Tariff", "Modifier", "Qty", "Fees"}
    if not expected_columns.issubset(df.columns):
        st.error("Charge sheet must contain: Description, Tariff, Modifier, Qty, Fees")
        st.stop()

    selected = st.multiselect(
        "Choose Procedures",
        df["Description"].tolist()
    )

    selected_rows = df[df["Description"].isin(selected)]
else:
    selected_rows = pd.DataFrame()

# -------------------------------------------------------
# FUNCTION: Fill the Excel template
# -------------------------------------------------------
def fill_excel_template(template_file, patient, member, provider, date, rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    # ----------------------------------------------
    # Fill header fields (Top section)
    # ----------------------------------------------
    ws["B4"] = patient
    ws["B5"] = member
    ws["B6"] = provider
    ws["B7"] = str(date)

    # ----------------------------------------------
    # BLUE LINE stays untouched (ROW 22)
    # Data must start from ROW 23 downwards
    # ----------------------------------------------
    START_ROW = 23

    current_row = START_ROW

    for _, r in rows.iterrows():
        ws.cell(row=current_row, column=1, value=r["Description"])
        ws.cell(row=current_row, column=2, value=r["Tariff"])
        ws.cell(row=current_row, column=3, value=r["Modifier"] if not pd.isna(r["Modifier"]) else "")
        ws.cell(row=current_row, column=4, value=r["Qty"])
        ws.cell(row=current_row, column=7, value=r["Fees"])  # FEES goes to column G

        # Format alignment
        for col in [1, 2, 3, 4, 7]:
            ws.cell(row=current_row, column=col).alignment = Alignment(horizontal="left")

        current_row += 1

    # ----------------------------------------------
    # TOTAL remains in G22 EXACTLY as your template
    # Template must already contain formula:
    # =SUM(G23:G200)
    # ----------------------------------------------

    # DO NOT touch G22
    # DO NOT write any totals here

    # Save output to memory
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# -------------------------------------------------------
# Generate & Download Output
# -------------------------------------------------------
if st.button("Generate Quotation"):
    if uploaded_template is None:
        st.error("Please upload quotation template.")
        st.stop()

    if selected_rows.empty:
        st.error("No items selected.")
        st.stop()

    output = fill_excel_template(
        uploaded_template,
        patient,
        member,
        provider,
        quote_date,
        selected_rows
    )

    st.success("Quotation generated successfully.")

    st.download_button(
        label="Download Quotation",
        data=output,
        file_name="quotation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
