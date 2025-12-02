import streamlit as st
import pandas as pd
import openpyxl

st.title("Radiology Quotation Auto-Generator")

st.write("""
Upload your quotation template and the charge sheet.
Enter the patient details and scan type, and the app will automatically
fill the quotation using the correct tab and CIMAS USD price column.
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
scan_type = st.text_input("Scan Type (Tab Name exactly e.g. 'XRAY RECIPES', 'CT SCAN RECIPES')")

# If charge sheet uploaded, list the available tabs
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

    # Validate inputs
    if not template_file:
        st.error("Upload quotation template file first.")
        st.stop()

    if not charge_sheet_file:
        st.error("Upload charge sheet file first.")
        st.stop()

    if scan_type.strip() == "":
        st.error("Enter scan type (must match a sheet/tab name).")
        st.stop()

    try:
        # ---------------------------------------------------
        # STEP 1: LOAD TEMPLATE
        # ---------------------------------------------------
        template_df = pd.read_excel(template_file)

        # ---------------------------------------------------
        # STEP 2: MATCH THE CORRECT TAB
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
                Available tabs: {xl.sheet_names}
            """)
            st.stop()

        # Load the selected sheet
        charge_df = pd.read_excel(charge_sheet_file, sheet_name=sheet_match)

        # ---------------------------------------------------
        # STEP 3: VALIDATE COLUMNS
        # ---------------------------------------------------
        if "Tariff" not in template_df.columns:
            st.error("Quotation template must have a column named 'Tariff'.")
            st.stop()

        if "TARIFF" not in charge_df.columns and "Tariff" not in charge_df.columns:
            st.error("Charge sheet must contain a 'TARIFF' column.")
            st.stop()

        # Fix charge sheet tariff column name
        if "TARIFF" in charge_df.columns:
            charge_df.rename(columns={"TARIFF": "Tariff"}, inplace=True)

        # Check CIMAS USD
        if "CIMAS USD" not in charge_df.columns:
            st.error("Charge sheet must contain a column named 'CIMAS USD'.")
            st.stop()

        # ---------------------------------------------------
        # STEP 4: MATCH TARIFF & ASSIGN PRICE
        # ---------------------------------------------------
        output_df = template_df.copy()

        for index, row in output_df.iterrows():
            tariff_code = row["Tariff"]

            match = charge_df[charge_df["Tariff"] == tariff_code]

            if not match.empty:
                price = match["CIMAS USD"].values[0]
                output_df.loc[index, "Price"] = price
            else:
                output_df.loc[index, "Price"] = None

        # ---------------------------------------------------
        # STEP 5: ADD PATIENT DETAILS
        # ---------------------------------------------------
        output_df["PatientName"] = patient_name
        output_df["MedicalAid"] = medical_aid

        # ---------------------------------------------------
        # STEP 6: EXPORT FINAL EXCEL
        # ---------------------------------------------------
        output_path = "Generated_Quotation.xlsx"
        output_df.to_excel(output_path, index=False)

        # ---------------------------------------------------
        # DOWNLOAD BUTTON
        # ---------------------------------------------------
        st.success("Quotation generated successfully!")

        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Final Quotation",
                data=f,
                file_name="Quotation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {str(e)}")
