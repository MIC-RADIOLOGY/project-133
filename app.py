import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# -----------------------------
# LOAD CHARGE SHEET
# -----------------------------
def load_charge_sheet(file):
    df = pd.read_excel(file)

    # Clean columns if needed
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Only keep rows with tariff numbers (valid scan lines)
    valid_rows = df[df["TARRIF"].notna()]

    # Ensure proper types
    valid_rows["EXAMINATION"] = valid_rows["EXAMINATION"].astype(str)
    valid_rows["TARRIF"] = valid_rows["TARRIF"].astype(str)
    valid_rows["MODIFIER"] = valid_rows["MODIFIER"].astype(str)
    valid_rows["QUANTITY"] = valid_rows["QUANTITY"]
    valid_rows["AMOUNT"] = valid_rows["AMOUNT"]

    return valid_rows


# -----------------------------
# PDF GENERATOR
# -----------------------------
def generate_pdf(patient_name, med_number, provider, scan_name, scan_row):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "MEDICAL QUOTATION", ln=True, align="C")

    pdf.ln(5)

    pdf.set_font("Arial", size=12)

    # Patient details
    pdf.cell(0, 8, f"Patient Name: {patient_name}", ln=True)
    pdf.cell(0, 8, f"Medical Aid Number: {med_number}", ln=True)
    pdf.cell(0, 8, f"Provider: {provider}", ln=True)

    pdf.ln(10)

    # Quotation table header
    pdf.set_font("Arial", "B", 12)
    pdf.cell(80, 8, "Examination", border=1)
    pdf.cell(30, 8, "Tariff", border=1, align="C")
    pdf.cell(25, 8, "Modifier", border=1, align="C")
    pdf.cell(25, 8, "Qty", border=1, align="C")
    pdf.cell(30, 8, "Amount", border=1, align="R")
    pdf.ln()

    # Data row
    pdf.set_font("Arial", size=12)
    pdf.cell(80, 8, scan_name, border=1)
    pdf.cell(30, 8, scan_row["TARRIF"], border=1, align="C")
    pdf.cell(25, 8, scan_row["MODIFIER"], border=1, align="C")
    pdf.cell(25, 8, str(scan_row["QUANTITY"]), border=1, align="C")
    pdf.cell(30, 8, f"${scan_row['AMOUNT']}", border=1, align="R")

    pdf.ln(15)

    # Total
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, f"Total Amount: ${scan_row['AMOUNT']}", ln=True)

    buffer = io.BytesIO()
    pdf.output(buffer)
    buffer.seek(0)
    return buffer


# -----------------------------
# UI LAYOUT
# -----------------------------
st.title("📄 Medical Quotation Generator")
st.write("Upload your charge sheet & fill in patient information to create a quotation.")

uploaded_file = st.file_uploader("Upload the Excel Charge Sheet", type=["xlsx"])

if uploaded_file:
    df = load_charge_sheet(uploaded_file)

    st.success("Charge sheet loaded successfully!")

    st.subheader("Step 1: Enter Patient Details")
    col1, col2 = st.columns(2)

    with col1:
        patient_name = st.text_input("Patient Name")
        med_number = st.text_input("Medical Aid Number")

    with col2:
        provider = st.text_input("Medical Aid Provider", value="CIMAS")

    st.subheader("Step 2: Select Scan")

    scan_list = df["EXAMINATION"].unique().tolist()
    selected_scan = st.selectbox("Choose Scan", scan_list)

    if selected_scan:
        scan_row = df[df["EXAMINATION"] == selected_scan].iloc[0]

        st.write("### Scan Details")
        st.write(f"**Tariff:** {scan_row['TARRIF']}")
        st.write(f"**Modifier:** {scan_row['MODIFIER']}")
        st.write(f"**Quantity:** {scan_row['QUANTITY']}")
        st.write(f"**Amount:** ${scan_row['AMOUNT']}")

        if st.button("Generate Quotation PDF"):
            pdf_file = generate_pdf(
                patient_name,
                med_number,
                provider,
                selected_scan,
                scan_row
            )

            st.success("Quotation generated!")

            st.download_button(
                label="Download PDF",
                data=pdf_file,
                file_name=f"quotation_{patient_name}.pdf",
                mime="application/pdf"
            )

else:
    st.info("Upload your charge sheet to begin.")
