import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Radiology Quotation Auto-Generator")

st.write("""
Upload your quotation template and the charge sheet.
Enter the patient details and scan type, and the app will automatically
fill the quotation using the correct tab and CIMAS USD price column.
""")

# -------------------------
# FILE UPLOAD
# -------------------------
template_file = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])
charge_sheet_file = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])

# -------------------------
# PATIENT & SCAN INPUT
# -------------------------
patient_name = st.text_input("Patient Full Name")
medical_aid = st.text_input("Medical Aid Number")
scan_type = st.text_input("Scan Type (e.g. 'Chest X-Ray', 'Head CT')")

# -------------------------
# GENERATE QUOTATION
# -------------------------
if st.button("Generate Quotation"):

    # Validate files
    if not template_file or not charge_sheet_file:
        st.error("Please upload both quotation template and charge sheet.")
        st.stop()

    if scan_type.strip() == "":
        st.error("Enter scan type.")
        st.stop()

    try:
        # -------------------------
        # LOAD TEMPLATE (HEADERS ON ROW 4 → header=3)
        # -------------------------
        template_df = pd.read_excel(template_file, header=3)
        template_df.columns = template_df.columns.str.strip().str.title()
        st.write("Template Columns Found:", template_df.columns.tolist())

        if "Tariff" not in template_df.columns:
            st.error("Quotation template must have a column named 'Tariff'.")
            st.stop()

        # -------------------------
        # LOAD CHARGE SHEET
        # -------------------------
        xl = pd.ExcelFile(charge_sheet_file)
        available_tabs = xl.sheet_names
        st.info(f"Available Tabs: {available_tabs}")

        # Map scan category to tab automatically
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
        charge_df.columns = charge_df.columns.str.strip().str.title()

        # -------------------------
        # VALIDATE CHARGE SHEET COLUMNS
        # -------------------------
        if "Tariff" not in charge_df.columns:
            st.error(f"'Tariff' column not found in tab '{sheet_name}'.")
            st.stop()
        if "Cimas Usd" not in charge_df.columns:
            st.error(f"'CIMAS USD' column not found in tab '{sheet_name}'.")
            st.stop()

        # Find scan/procedure column
        scan_column = None
        for col in charge_df.columns:
            if "scan" in col.lower() or "procedure" in col.lower():
                scan_column = col
                break
        if scan_column is None:
            st.error(f"No column for scan/procedure found in tab '{sheet_name}'.")
            st.stop()

        # -------------------------
        # FILTER FOR SPECIFIC SCAN
        # -------------------------
        filtered_charge_df = charge_df[charge_df[scan_column].str.contains(scan_type, case=False, na=False)]
        if filtered_charge_df.empty:
            st.warning(f"No matching scan found for '{scan_type}' in tab '{sheet_name}'. Using full tab.")
            filtered_charge_df = charge_df.copy()

        # -------------------------
        # MATCH TARIFFS AND ASSIGN PRICES
        # -------------------------
        output_df = template_df.copy()
        output_df["Price"] = None

        for idx, row in output_df.iterrows():
            tariff_code = row["Tariff"]
            match = filtered_charge_df[filtered_charge_df["Tariff"] == tariff_code]
            if not match.empty:
                output_df.at[idx, "Price"] = match["Cimas Usd"].values[0]

        missing_tariffs = output_df[output_df["Price"].isna()]["Tariff"].tolist()
        if missing_tariffs:
            st.warning(f"The following tariffs were not found: {missing_tariffs}")

        # -------------------------
        # ADD PATIENT DETAILS
        # -------------------------
        output_df["PatientName"] = patient_name
        output_df["MedicalAid"] = medical_aid

        # -------------------------
        # EXPORT TO EXCEL (IN-MEMORY)
        # -------------------------
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
