import streamlit as st
import pandas as pd
from io import BytesIO

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
scan_type = st.text_input("Scan Type (e.g. 'Chest X-Ray', 'Head CT')")

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
        st.error("Enter scan type.")
        st.stop()

    try:
        # ---------------------------------------------------
        # STEP 1: LOAD TEMPLATE
        # ---------------------------------------------------
        template_df = pd.read_excel(template_file)
        template_df.columns = template_df.columns.str.strip()  # clean column names

        if "Tariff" not in template_df.columns:
            st.error("Quotation template must have a column named 'Tariff'.")
            st.stop()

        # ---------------------------------------------------
        # STEP 2: LOAD CHARGE SHEET
        # ---------------------------------------------------
        xl = pd.ExcelFile(charge_sheet_file)
        available_tabs = xl.sheet_names
        st.info(f"Available Tabs in Charge Sheet: {available_tabs}")

        # Map scan category to tab
        scan_category = scan_type.split()[0].upper()  # e.g., XRAY, CT, MRI
        tab_map = {
            "XRAY": [tab for tab in available_tabs if "XRAY" in tab.upper()],
            "CT": [tab for tab in available_tabs if "CT" in tab.upper()],
            "MRI": [tab for tab in available_tabs if "MRI" in tab.upper()],
        }

        matching_tabs = tab_map.get(scan_category)
        if not matching_tabs:
            st.error(f"No matching tab found for scan category '{scan_category}'.")
            st.stop()

        sheet_name = matching_tabs[0]  # pick first matching tab
        charge_df = pd.read_excel(charge_sheet_file, sheet_name=sheet_name)
        charge_df.columns = charge_df.columns.str.strip()  # clean column names

        # ---------------------------------------------------
        # STEP 3: VALIDATE COLUMNS IN CHARGE SHEET
        # ---------------------------------------------------
        if "Tariff" not in charge_df.columns:
            st.error(f"'Tariff' column not found in tab '{sheet_name}'.")
            st.stop()

        if "CIMAS USD" not in charge_df.columns:
            st.error(f"'CIMAS USD' column not found in tab '{sheet_name}'.")
            st.stop()

        # Optional: column for scan name in charge sheet
        scan_column = None
        for col in charge_df.columns:
            if "scan" in col.lower() or "procedure" in col.lower():
                scan_column = col
                break
        if scan_column is None:
            st.error(f"No column for scan/procedure name found in tab '{sheet_name}'.")
            st.stop()

        # ---------------------------------------------------
        # STEP 4: FILTER CHARGE SHEET FOR SPECIFIC SCAN
        # ---------------------------------------------------
        filtered_charge_df = charge_df[charge_df[scan_column].str.contains(scan_type, case=False, na=False)]

        if filtered_charge_df.empty:
            st.warning(f"No matching scan found in tab '{sheet_name}' for '{scan_type}'. Using full tab.")
            filtered_charge_df = charge_df.copy()

        # ---------------------------------------------------
        # STEP 5: MATCH TARIFFS AND ASSIGN PRICES
        # ---------------------------------------------------
        output_df = template_df.copy()
        output_df["Price"] = None  # initialize

        for idx, row in output_df.iterrows():
            tariff_code = row["Tariff"]
            match = filtered_charge_df[filtered_charge_df["Tariff"] == tariff_code]
            if not match.empty:
                output_df.at[idx, "Price"] = match["CIMAS USD"].values[0]

        # Warn if some tariffs not found
        missing_tariffs = output_df[output_df["Price"].isna()]["Tariff"].tolist()
        if missing_tariffs:
            st.warning(f"The following tariffs were not found: {missing_tariffs}")

        # ---------------------------------------------------
        # STEP 6: ADD PATIENT DETAILS
        # ---------------------------------------------------
        output_df["PatientName"] = patient_name
        output_df["MedicalAid"] = medical_aid

        # ---------------------------------------------------
        # STEP 7: EXPORT TO EXCEL IN-MEMORY
        # ---------------------------------------------------
        output = BytesIO()
        output_df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.success("Quotation generated successfully!")
        st.download_button(
            label="Download Quotation",
            data=output,
            file_name="Quotation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {str(e)}")
