import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# -----------------------------------
# HELPER FUNCTION: Safe assignment to possibly merged cells
# -----------------------------------
def set_cell_value_safe(cell, value):
    """Assign value to cell, even if it's part of a merged range."""
    try:
        cell.value = value
    except AttributeError:
        # It's a merged cell, assign to top-left cell of the merged range
        for merged_range in cell.parent.merged_cells.ranges:
            if cell.coordinate in merged_range:
                cell.parent[merged_range.coord.split(":")[0]].value = value
                return

# -----------------------------------
# LOAD CHARGE SHEET WITH CATEGORIES
# -----------------------------------
def load_charge_sheet(file):
    df = pd.read_excel(file, header=None)
    df.columns = ["EXAMINATION", "TARRIF", "MODIFIER", "QUANTITY", "AMOUNT"]

    # Detect categories
    category = None
    categories = []
    for i, row in df.iterrows():
        if pd.isna(row["TARRIF"]) or row["TARRIF"] == "":
            # Main heading row
            category = str(row["EXAMINATION"])
            categories.append(category)
            df.at[i, "CATEGORY"] = None  # headings themselves have no category
        else:
            df.at[i, "CATEGORY"] = category

    # Keep only rows with tariffs
    df = df[df["TARRIF"].notna()]
    df["EXAMINATION"] = df["EXAMINATION"].astype(str)
    df["TARRIF"] = df["TARRIF"].astype(str)
    df["MODIFIER"] = df["MODIFIER"].astype(str)
    df["QUANTITY"] = pd.to_numeric(df["QUANTITY"], errors='coerce').fillna(0).astype(int)
    df["AMOUNT"] = pd.to_numeric(df["AMOUNT"], errors='coerce').fillna(0.0).astype(float)
    return df, categories

# -----------------------------------
# FILL EXCEL TEMPLATE AUTOMATICALLY
# -----------------------------------
def fill_excel_template(template_file, patient, member, provider, scan_row):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    patient_cell = member_cell = provider_cell = None
    scan_start_row = desc_col = tarif_col = modi_col = qty_col = amt_col = total_cell = None

    # Flexible detection of fields
    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                val = str(cell.value).strip().upper()
                if "PATIENT" in val and patient_cell is None:
                    patient_cell = ws.cell(row=cell.row, column=cell.column + 1)
                elif "MEMBER" in val and member_cell is None:
                    member_cell = ws.cell(row=cell.row, column=cell.column + 1)
                elif ("PROVIDER" in val or "EXAMINATION" in val) and provider_cell is None:
                    provider_cell = ws.cell(row=cell.row, column=cell.column + 1)
                elif "DESCRIPTION" in val and scan_start_row is None:
                    scan_start_row = cell.row + 1
                    desc_col = cell.column
                    tarif_col = cell.column + 1
                    modi_col = cell.column + 2
                    qty_col = cell.column + 3
                    amt_col = cell.column + 4
                elif "TOTAL" in val and total_cell is None:
                    total_cell = ws.cell(row=cell.row, column=cell.column + 6)

    # Warn user if fields missing
    missing_fields = []
    if not patient_cell: missing_fields.append("Patient Name")
    if not member_cell: missing_fields.append("Member Number")
    if not provider_cell: missing_fields.append("Provider")
    if not scan_start_row: missing_fields.append("Scan Table")
    if not total_cell: missing_fields.append("Total Cell")
    if missing_fields:
        st.warning(f"Could not detect these fields in template: {', '.join(missing_fields)}")

    # Assign values safely
    if patient_cell: set_cell_value_safe(patient_cell, patient)
    if member_cell: set_cell_value_safe(member_cell, member)
    if provider_cell: set_cell_value_safe(provider_cell, provider)

    if scan_start_row:
        set_cell_value_safe(ws.cell(row=scan_start_row, column=desc_col), scan_row["EXAMINATION"])
        set_cell_value_safe(ws.cell(row=scan_start_row, column=tarif_col), scan_row["TARRIF"])
        set_cell_value_safe(ws.cell(row=scan_start_row, column=modi_col), scan_row["MODIFIER"])
        set_cell_value_safe(ws.cell(row=scan_start_row, column=qty_col), int(scan_row["QUANTITY"]))
        set_cell_value_safe(ws.cell(row=scan_start_row, column=amt_col), float(scan_row["AMOUNT"]))

    if total_cell:
        set_cell_value_safe(total_cell, float(scan_row["AMOUNT"]))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -----------------------------------
# STREAMLIT INTERFACE
# -----------------------------------
st.title("📄 Medical Quotation Generator")

# Persistent session state
for key in ["charge_file", "template_file", "df", "categories", "patient_input", "member_input", "provider_input"]:
    if key not in st.session_state:
        st.session_state[key] = None if key in ["charge_file","template_file","df","categories"] else ""

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
if uploaded_charge is not None:
    st.session_state.charge_file = uploaded_charge

uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])
if uploaded_template is not None:
    st.session_state.template_file = uploaded_template

st.text_input("Patient Name", key="patient_input")
st.text_input("Medical Aid Number", key="member_input")
st.text_input("Medical Aid Provider", key="provider_input")

patient = st.session_state.patient_input
member = st.session_state.member_input
provider = st.session_state.provider_input

# -----------------------------
# Continue button loads charge sheet
# -----------------------------
if st.session_state.charge_file and st.session_state.template_file:
    if st.button("Continue") or st.session_state.df is None:
        st.session_state.df, st.session_state.categories = load_charge_sheet(st.session_state.charge_file)
        st.success("Charge sheet loaded!")

    df = st.session_state.df
    categories = st.session_state.categories

    if df is not None and categories:
        st.subheader("Select Scan Category")
        selected_category = st.selectbox("Category", categories)

        # Select scan under chosen category
        scan_list = df[df['CATEGORY'] == selected_category]['EXAMINATION'].tolist()
        selected_scan = st.selectbox("Select Scan", scan_list)

        if selected_scan:
            scan_row = df[(df['CATEGORY'] == selected_category) & (df['EXAMINATION'] == selected_scan)].iloc[0]

            st.write("### Scan Details:")
            st.write(f"**Tariff:** {scan_row['TARRIF']}")
            st.write(f"**Modifier:** {scan_row['MODIFIER']}")
            st.write(f"**Quantity:** {scan_row['QUANTITY']}")
            st.write(f"**Amount:** {scan_row['AMOUNT']}")

            if st.button("Generate Quotation"):
                output = fill_excel_template(
                    st.session_state.template_file,
                    patient,
                    member,
                    provider,
                    scan_row
                )
                if output:
                    st.success("Quotation generated!")
                    st.download_button(
                        "Download Excel Quotation",
                        data=output,
                        file_name=f"quotation_{patient}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
else:
    st.info("Upload your charge sheet and quotation template to continue.")
