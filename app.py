import os
import streamlit as st
import pandas as pd
from io import BytesIO

# ----------------------------------------
# DEFAULT FILE LOCATIONS (for fallback)
# ----------------------------------------
DATA_FOLDER = "data"
DEFAULT_CHARGE_SHEET = os.path.join(DATA_FOLDER, "charge_sheet.xlsx")
DEFAULT_TEMPLATE = os.path.join(DATA_FOLDER, "template.xlsx")

# ----------------------------------------
# LOAD FILES
# ----------------------------------------
def load_excel(path, label):
    """Load Excel or show a friendly error."""
    if not os.path.exists(path):
        st.error(f"❌ {label} not found. Make sure **{path}** exists in the repo.")
        return None

    try:
        return pd.read_excel(path)
    except Exception as e:
        st.error(f"❌ Failed to load {label}: {e}")
        return None

# ----------------------------------------
# STREAMLIT UI
# ----------------------------------------
st.set_page_config(page_title="Quotation Generator", layout="wide")
st.title("📄 Automated Quotation Generator")
st.write("Upload your files manually or use default charge sheet and template from `/data` folder.")

# ----------------------------------------
# FILE UPLOAD (manual)
# ----------------------------------------
st.subheader("📂 Upload Files (Optional)")

uploaded_charge_sheet = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

# Load files: uploaded takes priority, fallback to defaults
if uploaded_charge_sheet:
    try:
        charge_sheet = pd.read_excel(uploaded_charge_sheet)
        st.success("✅ Uploaded Charge Sheet loaded successfully!")
    except Exception as e:
        st.error(f"❌ Failed to read uploaded Charge Sheet: {e}")
        st.stop()
else:
    charge_sheet = load_excel(DEFAULT_CHARGE_SHEET, "Charge Sheet")

if uploaded_template:
    try:
        template = pd.read_excel(uploaded_template)
        st.success("✅ Uploaded Template loaded successfully!")
    except Exception as e:
        st.error(f"❌ Failed to read uploaded Template: {e}")
        st.stop()
else:
    template = load_excel(DEFAULT_TEMPLATE, "Quotation Template")

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
        selected_index = st.selectbox("Choose scan", matches.index)

        if selected_index is not None:
            selected_row = matches.loc[selected_index]

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
