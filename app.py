import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ------------------------------------------------------------
# FORCE WRITE INTO MERGED CELLS
# ------------------------------------------------------------
def write_force(ws, row, col, value):
    if col is None:
        return
    cell = ws.cell(row, col)
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            tl = mr.coord.split(":")[0]
            cell = ws[tl]
            break
    cell.value = value

# ------------------------------------------------------------
# PARSE UCS CHARGE SHEET
# ------------------------------------------------------------
def parse_charge_sheet(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)

    # Expected header aliases
    aliases = {
        "DESCRIPTION": ["DESCRIPTION", "EXAMINATION", "SCAN", "ITEM"],
        "TARIFF": ["TARRIF", "TARIFF"],
        "MODIFIER": ["MODIFIER", "MOD"],
        "QTY": ["QTY", "QUANTITY"],
        "FEES": ["FEES", "FEE", "AMOUNT"],   # Your sheet stores FEES inside AMOUNT
        "AMOUNT": ["AMOUNT", "TOTAL", "COST"]
    }

    header_row = None
    detected_cols = {}

    # Detect header row
    for i in range(len(df_raw)):
        row = [str(x).strip().upper() for x in df_raw.iloc[i].tolist()]
        for key, names in aliases.items():
            for n in names:
                if n in row:
                    detected_cols[key] = row.index(n)
        if len(detected_cols) >= 4:  # enough to treat as header
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not locate a valid header row.")

    # Load again using header
    df = pd.read_excel(uploaded_file, header=header_row, dtype=str)
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Map headers
    rename_map = {}
    for col in df.columns:
        u = col.strip().upper()
        if u in ["EXAMINATION", "DESCRIPTION"]:
            rename_map[col] = "DESCRIPTION"
        elif u in ["TARRIF", "TARIFF"]:
            rename_map[col] = "TARIFF"
        elif u in ["MOD", "MODIFIER"]:
            rename_map[col] = "MODIFIER"
        elif u in ["QTY", "QUANTITY"]:
            rename_map[col] = "QTY"
        elif u in ["FEES"]:
            rename_map[col] = "FEES"
        elif u in ["AMOUNT"]:
            rename_map[col] = "AMOUNT"

    df = df.rename(columns=rename_map)

    # Handle FEES missing = use AMOUNT
    if "FEES" not in df.columns and "AMOUNT" in df.columns:
        df["FEES"] = df["AMOUNT"]

    # Required
    required = ["DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"]
    for r in required:
        if r not in df.columns:
            df[r] = ""

    # Clean numeric
    def to_number(x):
        if pd.isna(x):
            return 0
        x = str(x).replace(",", "").replace("$", "").strip()
        try:
            return float(x)
        except:
            return 0

    df["QTY"] = df["QTY"].apply(lambda x: int(str(x).split('.')[0]) if str(x).replace('.', '').isdigit() else 1)
    df["FEES"] = df["FEES"].apply(to_number)
    df["AMOUNT"] = df["AMOUNT"].apply(to_number)

    # Build row dicts
    rows = []
    for _, r in df.iterrows():
        if str(r["DESCRIPTION"]).strip().upper() in ["TOTAL", ""]:
            continue
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
# FILL TEMPLATE
# ------------------------------------------------------------
def fill_excel_template(template_file, patient, member, provider, rows):
    workbook = openpyxl.load_workbook(template_file)
    ws = workbook.active

    colmap = {
        "DESCRIPTION": 1,
        "TARIFF": 2,
        "MODIFIER": 3,
        "QTY": 5,
        "FEES": 6,
        "AMOUNT": 7
    }

    write_force(ws, 13, 2, f"FOR PATIENT: {patient}")
    write_force(ws, 14, 2, f"MEMBER NUMBER: {member}")
    write_force(ws, 12, 5, provider)

    start_row = 20
    rowptr = start_row

    for r in rows:
        write_force(ws, rowptr, colmap["DESCRIPTION"], r["SCAN"])
        write_force(ws, rowptr, colmap["TARIFF"], r["TARIFF"])
        write_force(ws, rowptr, colmap["MODIFIER"], r["MODIFIER"])
        write_force(ws, rowptr, colmap["QTY"], r["QTY"])
        write_force(ws, rowptr, colmap["FEES"], r["FEES"])
        write_force(ws, rowptr, colmap["AMOUNT"], r["AMOUNT"])
        rowptr += 1

    total = sum([x["AMOUNT"] for x in rows])
    write_force(ws, 22, 7, total)

    out = io.BytesIO()
    workbook.save(out)
    out.seek(0)
    return out

# ------------------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

template = st.file_uploader("Upload Template", type=["xlsx"])
charge = st.file_uploader("Upload Charge Sheet", type=["xlsx"])
patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Provider / ATT")

if template and charge and patient and member and provider:
    try:
        rows = parse_charge_sheet(charge)
        if st.button("Generate"):
            output = fill_excel_template(template, patient, member, provider, rows)
            st.success("Quotation generated successfully.")
            st.download_button(
                "Download Quotation",
                output,
                file_name=f"quotation_{patient.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error: {e}")
