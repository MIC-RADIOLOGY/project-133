import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Scan Quotation Generator", layout="wide")

# ---------------------- Load and parse charge sheet ----------------------

@st.cache_data(show_spinner=False)
def load_charge_sheet(file) -> tuple:
    # Load charge sheet without header because your data is complex
    df = pd.read_excel(file, header=None)

    # Rename columns to expected names or placeholders for processing
    df.columns = ["EXAMINATION", "TARRIF", "MODIFIER", "QUANTITY", "AMOUNT"]

    # Clean data types and strip strings
    df["EXAMINATION"] = df["EXAMINATION"].astype(str).str.strip()
    df["TARRIF"] = df["TARRIF"].apply(lambda x: str(int(x)) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else None)
    df["MODIFIER"] = df["MODIFIER"].astype(str).str.strip()
    df["QUANTITY"] = pd.to_numeric(df["QUANTITY"], errors='coerce').fillna(0).astype(int)
    df["AMOUNT"] = pd.to_numeric(df["AMOUNT"], errors='coerce').fillna(0.0)

    # For hierarchical detection
    categories = []
    subcategories = []

    current_category = None
    current_subcategory = None

    exclude_scans = {"TOTAL", "CO - Payment", "FF"}

    # Add columns for category and subcategory
    df["CATEGORY"] = None
    df["SUBCATEGORY"] = None

    for idx, row in df.iterrows():
        ex_upper = row["EXAMINATION"].upper()

        # Detect main category: usually upper case, no tariff, no quantity/amount
        if pd.isna(row["TARRIF"]) and row["QUANTITY"] == 0 and row["AMOUNT"] == 0.0:
            if ex_upper in exclude_scans:
                # ignore these as category/subcategory
                current_subcategory = None
                continue

            # Heuristic: treat lines with all caps and length > 3 as category
            if ex_upper == row["EXAMINATION"] and len(ex_upper) > 3 and ex_upper.isalpha() or " " in ex_upper:
                # It's a main category
                current_category = row["EXAMINATION"]
                categories.append(current_category)
                current_subcategory = None
            else:
                # Treat as subcategory (if not main category)
                current_subcategory = row["EXAMINATION"]
                if current_subcategory not in subcategories:
                    subcategories.append(current_subcategory)
        else:
            # Assign category and subcategory for scans only if not excluded
            if ex_upper not in exclude_scans:
                df.at[idx, "CATEGORY"] = current_category
                df.at[idx, "SUBCATEGORY"] = current_subcategory

    # Filter dataframe to only rows with category assigned and exclude rows like TOTAL, FF, CO-Payment
    df_scans = df[df["CATEGORY"].notna()]
    df_scans = df_scans[~df_scans["EXAMINATION"].str.upper().isin(exclude_scans)]

    return df_scans, categories

# ----------------------- Load and fill quotation template ------------------

def fill_excel_template(template_file, patient_name, medical_aid_number, medical_aid_provider, selected_scans):
    # Load the template into openpyxl workbook
    wb = load_workbook(template_file)
    ws = wb.active

    # Simple heuristics to find cells for patient data (you should adapt these for your template):
    # Search cells containing specific keywords, then fill next cell (example approach)
    for row in ws.iter_rows(min_row=1, max_row=20, min_col=1, max_col=10):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                val = cell.value.lower()
                if "for patient" in val or "patient" in val:
                    ws.cell(row=cell.row, column=cell.column + 1, value=patient_name)
                elif "member number" in val or "medical aid number" in val:
                    ws.cell(row=cell.row, column=cell.column + 1, value=medical_aid_number)
                elif "medical examination" in val or "medical aid provider" in val:
                    ws.cell(row=cell.row, column=cell.column + 1, value=medical_aid_provider)

    # Start inserting scans from row after header - you will need to adjust based on your template layout
    start_row = 23
    current_row = start_row

    for scan in selected_scans:
        ws.cell(row=current_row, column=1, value=scan["EXAMINATION"])
        ws.cell(row=current_row, column=2, value=int(scan["TARRIF"]))
        ws.cell(row=current_row, column=3, value=scan["MODIFIER"] if scan["MODIFIER"] != "nan" else "")
        ws.cell(row=current_row, column=4, value=scan["QUANTITY"])
        ws.cell(row=current_row, column=5, value=scan["AMOUNT"])
        current_row += 1

    # Save to BytesIO buffer
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ----------------------------- Streamlit app -----------------------------

st.title("Scan Quotation Generator")

# Upload charge sheet
charge_file = st.file_uploader("Upload Charge Sheet Excel (Charge Sheet format)", type=["xlsx"])

if charge_file:
    with st.spinner("Loading charge sheet..."):
        try:
            df_charge, categories = load_charge_sheet(charge_file)
            st.success("Charge sheet loaded successfully!")
        except Exception as e:
            st.error(f"Failed to load charge sheet: {e}")
            st.stop()

    # Select main category
    selected_category = st.selectbox("Select Scan Main Category", options=categories)

    # Filter subcategories for selected category
    df_subcats = df_charge[df_charge["CATEGORY"] == selected_category]
    subcategories = df_subcats["SUBCATEGORY"].dropna().unique().tolist()
    selected_subcategory = st.selectbox("Select Subcategory", options=subcategories)

    # Filter scans for selected category and subcategory
    df_scans = df_subcats[df_subcats["SUBCATEGORY"] == selected_subcategory]

    # Select scan(s)
    scan_options = df_scans["EXAMINATION"].tolist()
    selected_scan_names = st.multiselect("Select Scans", options=scan_options)

    # Filter selected scan details
    selected_scans = df_scans[df_scans["EXAMINATION"].isin(selected_scan_names)].to_dict('records')

    # Patient info input
    st.markdown("### Patient Information")
    patient_name = st.text_input("Patient Name")
    medical_aid_number = st.text_input("Medical Aid Number")
    medical_aid_provider = st.text_input("Medical Aid Provider", value="CIMAS")

    # Upload quotation template
    template_file = st.file_uploader("Upload Quotation Template Excel", type=["xlsx"])

    # Continue button
    if st.button("Continue"):
        if not patient_name or not medical_aid_number or not selected_scans or not template_file:
            st.error("Please fill all patient info, select scans, and upload the quotation template.")
        else:
            with st.spinner("Generating quotation..."):
                output_excel = fill_excel_template(
                    template_file,
                    patient_name,
                    medical_aid_number,
                    medical_aid_provider,
                    selected_scans
                )
                st.success("Quotation generated successfully!")

                st.download_button(
                    label="Download Filled Quotation",
                    data=output_excel,
                    file_name="filled_quotation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
