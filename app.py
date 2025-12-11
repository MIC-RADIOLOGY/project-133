import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ------------------------------------------------------------
# NORMALIZE HEADERS
# ------------------------------------------------------------
def normalize_header(h):
    if not isinstance(h, str):
        return h
    h = h.strip().upper()

    replacements = {
        "TARRIF": "TARIFF",
        "MOD": "MODIFIER",
        "DESCRIPTION": "DESCRIPTION",
        "QTY": "QTY",
        "FEES": "FEES",
        "AMOUNT": "AMOUNT",
    }
    return replacements.get(h, h)


# ------------------------------------------------------------
# SAFE WRITE FOR MERGED CELLS
# ------------------------------------------------------------
def write_force(ws, row, col, value):
    if col is None:
        return
    cell = ws.cell(row, col)

    # If cell is inside merged region, write to top-left instead
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            tl = mr.coord.split(":")[0]
            cell = ws[tl]
            break

    cell.value = value


# ------------------------------------------------------------
# PARSE CHARGE SHEET
# ------------------------------------------------------------
def parse_charge_sheet(uploaded_file):
    df = pd.read_excel(uploaded_file, dtype=str)
    df.columns = [normalize_header(c) for c in df.columns]

    required = {"DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Convert quantities
    df["QTY"] = df["QTY"].apply(lambda x: int(x) if str(x).isdigit() else 1)

    # FEES used for print
    df["FEES"] = df["FEES"].apply(
        lambda x: float(str(x).replace("$", "").replace(",", "")) if pd.notna(x) else 0.0
    )

    # AMOUNT used for total
    df["AMOUNT"] = df["AMOUNT"].apply(
        lambda x: float(str(x).replace("$", "").replace(",", "")) if pd.notna(x) else 0.0
    )

    rows = []
    for _, r in df.iterrows():
        rows.append({
            "SCAN": r["DESCRIPTION"],
            "TARIFF": r["TARIFF"],
            "MODIFIER": r["MODIFIER"],
            "QTY": r["QTY"],
            "FEES": r["FEES"],
            "AMOUNT": r["AMOUNT"]
        })
    return rows


# ------------------------------------------------------------
# FILL EXCEL TEMPLATE
# ------------------------------------------------------------
def fill_excel_template(template_file, patient, member, provider, rows):
    workbook = openpyxl.load_workbook(template_file)
    ws = workbook.active

    # Column mapping (your FINAL mapping)
    cols = {
        "DESCRIPTION": 1,  # A
        "TARIFF": 2,       # B
        "MODIFIER": 3,     # C
        "QTY": 5,          # E
        "FEES": 6,         # F
        "AMOUNT": 7        # G
    }

    # Header information in template
    write_force(ws, 13, 2, f"FOR PATIENT: {patient}")
    write_force(ws, 14, 2, f"MEMBER NUMBER: {member}")
    write_force(ws, 12, 5, provider)

    # Write items starting at row 20
    start_row = 20
    rowptr = start_row

    for item in rows:
        write_force(ws, rowptr, cols["DESCRIPTION"], item["SCAN"])
        write_force(ws, rowptr, cols["TARIFF"], item["TARIFF"])
        write_force(ws, rowptr, cols["MODIFIER"], item["MODIFIER"])
        write_force(ws, rowptr, cols["QTY"], item["QTY"])
        write_force(ws, rowptr, cols["FEES"], item["FEES"])
        rowptr += 1

    # Total in G22 using AMOUNT
    total_amount = sum([r["AMOUNT"] for r in rows])
    write_force(ws, 22, 7, total_amount)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


# ------------------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

template_file = st.file_uploader("Upload Excel Template", type=["xlsx"])
charge_sheet_file = st.file_uploader("Upload Charge Sheet", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Provider / ATT")

if template_file and charge_sheet_file and patient and member and provider:
    try:
        rows = parse_charge_sheet(charge_sheet_file)

        if st.button("Generate Quotation"):
            output = fill_excel_template(template_file, patient, member, provider, rows)
            st.success("Quotation generated successfully.")

            st.download_button(
                label="Download Quotation",
                data=output,
                file_name=f"quotation_{patient.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
