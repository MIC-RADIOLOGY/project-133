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
# File Uploaders
# ---------------------------------------------------
template_file = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])
charge_sheet_file = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])

# ---------------------------------------------------
# Inputs
# ---------------------------------------------------
patient_name = st.text_input("Patient Full Name")
medical_aid = st.text_input("Medical Aid Number")
scan_type = st.text_input("Scan Type (Tab Name in Charge Sheet e.g 'USS', 'XRAY', 'CT')")

# ---------------------------------------------------
# Processing
# ---------------------------------------------------
if st.button("Generate Quotation"):

    if not template_file:
        st.error("Please upload a quotation template Excel file.")
    elif not charge_sheet_file:
        st.error("Please upload a charge sheet Excel file.")
    elif scan_type == "":
        st.error("Please enter the scan type (this must match the tab name).")
    else:
        try:
            # Load template (as normal dataframe)
            template_df = pd.read_excel(template_file)

            # Load correct tab from charge sheet
            charge_df = pd.read_excel(charge_sheet_file, sheet_name=scan_type)

            # Assume template has a column "Tariff"
            # And charge sheet has columns "Tariff" and "Price"

            output_df = template_df.copy()

            for index, row in output_df.iterrows():
                tariff_code = row.get("Tariff")

                match = charge_df[charge_df["Tariff"] == tariff_code]

                if not match.empty:
                    output_df.loc[index, "Price"] = match["Price"].values[0]
                else:
                    output_df.loc[index, "Price"] = None

            # Add patient info
            output_df["PatientName"] = patient_name
            output_df["MedicalAid"] = medical_aid

            # Save output
            output_path = "Generated_Quotation.xlsx"
            output_df.to_excel(output_path, index=False)

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
