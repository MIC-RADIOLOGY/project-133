import os
import streamlit as st
import pandas as pd

# ----------------------------------------
# DEFAULT FILE LOCATIONS (Cloud + Local)
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
st.write("Using **default charge sheet and template** from `/data` folder.")

# Load charge sheet
charge_sheet = load_excel(DEFAULT_CHARGE_SHEET, "Charge Sheet")

# Load quotation template
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
        selected = st.selectbox("Choose scan", matches.index)

        if selected:
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
            filled.to_excel(output_path, index=False)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="⬇ Download Quotation",
                    data=f,
                    file_name="quotation_generated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.success("Quotation generated successfully!")
