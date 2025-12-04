import streamlit as st
import pandas as pd
from io import BytesIO

# ----------------------------------------
# LOAD FILES (manual input only)
# ----------------------------------------
def load_excel(uploaded_file, label):
    """Load Excel or show a friendly error."""
    if uploaded_file is None:
        st.error(f"❌ {label} not uploaded. Please upload the file to continue.")
        return None

    try:
        return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ Failed to load {label}: {e}")
        return None


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

# Load files
charge_sheet = load_excel(charge_sheet_file, "Charge Sheet")
template = load_excel(template_file, "Quotation Template")

if charge_sheet is None or template is None:
    st.stop()

# ----------------------------------------
# INPUT SECTION
# ----------------------------------------
st.subheader("🔍 Find Scan")
scan_name = st.text_input("Enter scan name (e.g., CT Chest, MRI Brain, Ultrasound Whole Abdomen):")

if scan_name:
    matches = charge_sheet[
        charge_sheet.apply(lambda row: row.astype(str).str.contains(scan_name, case=False, na=False).any(), axis=1)
    ]

    if matches.empty:
        st.warning("⚠ No matching scan found.")
    else:
        st.success(f"Found **{len(matches)}** matching scan(s).")
        st.dataframe(matches)

        # Select one to use for quotation
        selected = st.selectbox("Choose scan", matches.index)

        if selected is not None:
            selected_row = matches.loc[selected]

            # ----------------------------------------
            # Insert into template
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
            output_path = "generated_quotation.xlsx"
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
