import streamlit as st
import pandas as pd
import openpyxl

st.title("Radiology Quotation Auto-Generator")

st.write("""
Upload your quotation template and the charge sheet.
Enter the patient details and scan type, and the app will automatically
fill the quotation using the correct tab and tariffs.
""")

# ---------------------------------------------------
# FILE UPLOAD SECTION
# ---------------------------------------------------
template_file = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])
charge_sheet_file = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])

# ---------------------------------------------------
# INPUT SECTION
# ---------------------------------------------------
patient_name = st.text_input("Patient Full Name")
medical_aid = st.text_input("Medical Aid Number")
scan_type = st.text_input("Scan Type (MUST match tab name, e.g. 'USS', 'XRAY', 'CT CHEST')")

# If a charge sheet is uploaded, show its tabs
if charge_sheet_file:
    try:
        xl = pd.ExcelFile(charge_sheet_file)
        st.info(f"Available Tabs in Charge Sheet: {xl.sheet_names}")
    except:
        st.warning("Could not read sheet names.")


# ---------------------------------------------------
# PROCESS BUTTON
# ---------------------------------------------------
if st.button("Generate Quotation"):

    # 1. Validate file uploads & fields
    if not template_file:
        st.error("Please upload a quotation template file.")
        st.stop()

    if not charge_sheet_file:
        st.error("Please upload a charge sheet file.")
        st.stop()

    if scan_type.strip() == "":
        st.error("Please enter the scan type (matching tab name).")
        st.stop()

    try:
        # ---------------------------------------------------
        # STEP 1: LOAD TEMPLATE
        # ---------------------------------------------------
        template_df = pd.read_excel(template_file)

        # ---------------------------------------------------
        # STEP 2: FIND CORRECT TAB (MATCH EVEN IF DIFFERENT)
        # ---------------------------------------------------
        xl = pd.ExcelFile(charge_sheet_file)
        normalized_target = scan_type.strip().lower()
        sheet_match = None

        for sheet in xl.sheet_names:
            if sheet.strip().lower() == normalized_target:
                sheet_match = sheet
                break

        if sheet_match is None:
            st.error(f"""
                Could not find a matching tab for '{scan_type}'.
                Available tabs are: {xl.sheet_names}
            """)
            st.stop()

        # Load matching tab
        charge_df = pd.read_excel(charge_sheet_file, sheet_name=sheet_match)

        # ---------------------------------------------------
        # STEP 3: MATCH TARIFFS & FILL PRICES
        # ---------------------------------------------------
        output_df = template_df.copy()

        if "Tariff" not in output_df.columns:
            st.error("Quotation template must have a 'Tariff' column.")
            st.stop()

        if "Tariff" not in charge_df.columns or "Price" not in charge_df.columns:
            st.error("Charge sheet tab must have 'Tariff' and 'Price' columns.")
            st.stop()

        for index, row in output_df.iterrows():
            tariff_code = row["Tariff"]

            match = charge_df[charge_df["Tariff"] == tariff_code]

            if not match.empty:
                output_df.loc[index, "Price"] = match["Price"].values[0]
            else:
                output_df.loc[index, "Price"] = None  # Tariff not found

        # ---------------------------------------------------
        # STEP 4: ADD PATIENT DETAILS
        # ---------------------------------------------------
        output_df["PatientName"] = patient_name
        output_df["MedicalAid"] = medical_aid

        # ---------------------------------------------------
        # STEP 5: GENERATE FINAL EXCEL
        # ---------------------------------------------------
        output_path = "Generated_Quotation.xlsx"
        output_df.to_excel(output_path, index=False)

        # ---------------------------------------------------
        # STEP 6: DOWNLOAD BUTTON
        # ---------------------------------------------------
        st.success("Quotation successfully generated!")

        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Final Quotation",
                data=f,
                file_name="Quotation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {str(e)}")
