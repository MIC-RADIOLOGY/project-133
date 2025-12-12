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
        "FEES": ["FEES", "FEE", "AMOUNT"],  # FEES stored inside AMOUNT in UCS
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
        if len(detected_cols) >= 4:  # enough columns found → valid header row
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not locate a valid header row.")

    # Load using the header row
    df = pd.read_excel(uploaded_file, header=header_row, dtype=str)
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Map headers to standard names
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
        elif u == "FEES":
            rename_map[col] = "FEES"
        elif u == "AMOUNT":
            rename_map[col] = "AMOUNT"

    df = df.rename(columns=rename_map)

    # FEES missing? → Use AMOUNT column
    if "FEES" not in df.columns and "AMOUNT" in df.columns:
        df["FEES"] = df["AMOUNT"]

    # Ensure all required columns exist
    required = ["DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"]
    for r in required:
        if r not in df.columns:
            df[r] = ""

    # Numeric cleaner
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
        "DESCRIPTION": 1,  # A
        "TARIFF": 2,       # B
        "MODIFIER": 3,     # C
        "QTY": 5,          # E
        "FEES": 6,         # F
        "AMOUNT": 7        # G
    }

    # Header info
    write_force(ws, 13, 2, f"FOR PATIENT: {patient}")
    write_force(ws, 14, 2, f"MEMBER NUMBER: {member}")
    write_force(ws, 12, 5, provider)

    # Item rows start at row 20
    rowptr = 20

    for r in rows:
        write_force(ws, rowptr, colmap["DESCRIPTION"], r["SCAN"])
        write_force(ws, rowptr, colmap["TARIFF"], r["TARIFF"])
        write_force(ws, rowptr, colmap["MODIFIER"], r["MODIFIER"])
        write_force(ws, rowptr, colmap["QTY"], r["QTY"])
        write_force(ws, rowptr, colmap["FEES"], r["FEES"])
        write_force(ws, rowptr, colmap["AMOUNT"], r["AMOUNT"])
        rowptr += 1

    # Total in G22
    total = sum([x["AMOUNT"] for x in rows])
    write_force(ws, 22, 7, total)

    out = io.BytesIO()
    workbook.save(out)
    out.seek(0)
    return out

# ------------------------------------------------------------
# STREAMLIT UI WITH PARSE + GENERATE BUTTONS
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

template = st.file_uploader("Upload Excel Template", type=["xlsx"])
charge = st.file_uploader("Upload Charge Sheet", type=["xlsx"])
patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Provider / ATT")

# Session storage for parsed data
if "parsed_rows" not in st.session_state:
    st.session_state.parsed_rows = None

# -----------------------
# PARSE BUTTON
# -----------------------
if charge and st.button("Parse Charge Sheet"):
    try:
        rows = parse_charge_sheet(charge)
        st.session_state.parsed_rows = rows
        st.success("Charge sheet parsed successfully.")
        st.dataframe(pd.DataFrame(rows))
    except Exception as e:
        st.error(f"Error: {e}")

# Show parsed data if already parsed
if st.session_state.parsed_rows is not None:
    st.subheader("Parsed Charge Sheet")
    st.dataframe(pd.DataFrame(st.session_state.parsed_rows))

# -----------------------
# GENERATE BUTTON
# -----------------------
if template and st.session_state.parsed_rows and patient and member and provider:
    if st.button("Generate Quotation"):
        try:
            output = fill_excel_template(template, patient, member, provider, st.session_state.parsed_rows)
            st.success("Quotation generated successfully.")
            st.download_button(
                "Download Quotation",
                output,
                file_name=f"quotation_{patient.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")
