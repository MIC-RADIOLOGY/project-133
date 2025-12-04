import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# -----------------------------------
# LOAD CHARGE SHEET
# -----------------------------------
def load_charge_sheet(file):
    df = pd.read_excel(file)

    # Normalize column names
    df.columns = [c.strip().upper() for c in df.columns]

    # Keep only valid scan rows
    df = df[df["TARRIF"].notna()]

    df["EXAMINATION"] = df["EXAMINATION"].astype(str)
    df["TARRIF"] = df["TARRIF"].astype(str)
    df["MODIFIER"] = df["MODIFIER"].astype(str)

    return df

# -----------------------------------
# FILL DOCX TEMPLATE
# -----------------------------------
def fill_template(template_file, replacements):
    doc = Document(template_file)

    # Replace inside paragraphs
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

    # Replace inside tables as well
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, val)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# ===================================
# STREAMLIT APP USER INTERFACE
# ===================================
st.title("📄 Medical Quotation Generator (With Template Upload)")

charge_file = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
template_file = st.file_uploader("Upload Quotation Template (DOCX)", type=["docx"])

if charge_file and template_file:

    df = load_charge_sheet(charge_file)
    st.success("Charge sheet loaded successfully!")

    st.subheader("Patient Information")
    col1, col2 = st.columns(2)

    with col1:
        patient_name = st.text_input("Patient Name")
        med_number = st.text_input("Medical Aid Number")

    with col2:
        provider = st.text_input("Medical Aid Provider", value="CIMAS")

    st.subheader("Select Scan")
    scan_list = df["EXAMINATION"].unique().tolist()

    selected_scan = st.selectbox("Choose Scan", scan_list)

    if selected_scan:
        scan_row = df[df["EXAMINATION"] == selected_scan].iloc[0]

        st.write("### Scan Details")
        st.write(f"**Tariff:** {scan_row['TARRIF']}")
        st.write(f"**Modifier:** {scan_row['MODIFIER']}")
        st.write(f"**Quantity:** {scan_row['QUANTITY']}")
        st.write(f"**Amount:** {scan_row['AMOUNT']}")

        if st.button("Generate Quotation"):

            # Build replacement dictionary
            replacements = {
                "{{PATIENT_NAME}}": patient_name,
                "{{MEDICAL_AID_NUMBER}}": med_number,
                "{{PROVIDER}}": provider,
                "{{SCAN_NAME}}": selected_scan,
                "{{TARRIF}}": scan_row["TARRIF"],
                "{{MODIFIER}}": scan_row["MODIFIER"],
                "{{QUANTITY}}": str(scan_row["QUANTITY"]),
                "{{AMOUNT}}": str(scan_row["AMOUNT"]),
                "{{TOTAL}}": str(scan_row["AMOUNT"])
            }

            output_docx = fill_template(template_file, replacements)

            st.success("Quotation generated!")

            st.download_button(
                "Download Quotation",
                data=output_docx,
                file_name=f"quotation_{patient_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

else:
    st.info("Please upload both the charge sheet and the template file.")
