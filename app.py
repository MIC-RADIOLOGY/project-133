import streamlit as st
import pandas as pd
from io import BytesIO

# ----------------------------------------
# STREAMLIT UI
# ----------------------------------------
st.set_page_config(page_title="Quotation Generator", layout="wide")
st.title("📄 Automated Quotation Generator")
st.write("Upload your **Charge Sheet** and **Quotation Template** to generate quotations.")

# ----------------------------------------
# FILE UPLOAD
# ----------------------------------------
st.subheader("📂 Upload Files")

charge_sheet_file = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
template_file = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

if charge_sheet_file and template_file:
    try:
        charge_sheet = pd.read_excel(charge_sheet_file)
        template = pd.read_excel(template_file)
    except Exception as e:
        st.error(f"❌ Failed to read uploaded files: {e}")
        st.stop()

    st.success("✅ Files uploaded successfully!")

    # ----------------------------------------
    # INPUT SECTION
    # ----------------------------------------
    st.subheader("🔍 Find Scan")
    scan_name = st.text_input("Enter scan name (e.g., CT Chest, MRI Brain, Ultrasound Whole Abdomen):")

    if scan_name:
        # Find matching scans
        matches = charge_sheet[
            charge_sheet.apply(lambda row: row.astype(str).str.contains(scan_name, case=False, na=False).any(), axis=1)
        ]

        if matches.empty:
            st.warning("⚠ No matching scan found.")
        else:
            st.success(f"Found **{len(matches)}** matching scan(s).")
            st.dataframe(matches)

            # Select one row to use for quotation
            selected_index = st.selectbox("Choose scan", matches.index)

            if selected_index is not None:
                selected_row = matches.loc[selected_index]

                # ----------------------------------------
                # GENERATE QUOTATION
                # ----------------------------------------
                st.subheader("📄 Generated Quotation")

                filled = template.copy()

                # Replace placeholders in template
                for column in filled.columns:
                    filled[column] = filled[column].astype(str)
                    filled[column] = filled[column].str.replace("{SCAN_NAME}", str(selected_row.get("Description", "")))
                    filled[column] = filled[column].str.replace("{PRICE}", str(selected_row.get("CIMAS USD", "")))

                st.dataframe(filled)

                # ----------------------------------------
                # DOWNLOAD BUTTON
                # ----------------------------------------
                output = BytesIO()
                filled.to_excel(output, index=False)
                output.seek(0)

                st.download_button(
                    label="⬇ Download Quotation",
                    data=output,
                    file_name="quotation_generated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.success("Quotation generated successfully!")
else:
    st.info("Please upload both Charge Sheet and Quotation Template to proceed.")
