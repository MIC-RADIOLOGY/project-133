import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------------------------
# Fixed scan-type list (Option 2)
# ---------------------------
SCAN_TYPES = [
    "ULTRA SOUND DOPPLERS",
    "ULTRA SOUND",
    "CT SCAN",
    "MRI",
    "X-RAY",
    "MAMMOGRAPHY"
]

# ------------------------------------------------------------
# Utility: safe write into merged cells
# ------------------------------------------------------------
def write_force(ws, row, col, value):
    if col is None:
        return
    cell = ws.cell(row, col)
    # If inside merged range, get top-left cell
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            tl = mr.coord.split(":")[0]
            cell = ws[tl]
            break
    cell.value = value

# ------------------------------------------------------------
# Parse charge sheet (robust header detection + normalization)
# ------------------------------------------------------------
def parse_charge_sheet(uploaded_file):
    """
    Returns list of dicts with keys:
    SCAN (DESCRIPTION), TARIFF, MODIFIER, QTY, FEES, AMOUNT
    """
    df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)

    # header aliases we will accept
    aliases = {
        "DESCRIPTION": ["DESCRIPTION", "EXAMINATION", "SCAN", "ITEM"],
        "TARIFF": ["TARRIF", "TARIFF"],
        "MODIFIER": ["MODIFIER", "MOD"],
        "QTY": ["QTY", "QUANTITY"],
        "FEES": ["FEES", "FEE"],
        "AMOUNT": ["AMOUNT", "AMT", "TOTAL", "COST"]
    }

    header_row = None
    # attempt to detect header row by scanning each row for at least 3-4 known headers
    for i in range(len(df_raw)):
        row = [str(x).strip().upper() if pd.notna(x) else "" for x in df_raw.iloc[i].tolist()]
        found = 0
        for name_list in aliases.values():
            for n in name_list:
                if n in row:
                    found += 1
                    break
        if found >= 3:
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not locate a header row. Ensure the charge sheet contains a header like EXAMINATION / TARRIF / MODIFIER / QUANTITY / AMOUNT")

    # read again with header_row
    df = pd.read_excel(uploaded_file, header=header_row, dtype=str)
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Build rename map
    rename_map = {}
    for col in df.columns:
        u = col.strip().upper()
        if u in ["EXAMINATION", "DESCRIPTION", "SCAN", "ITEM"]:
            rename_map[col] = "DESCRIPTION"
        elif u in ["TARRIF", "TARIFF"]:
            rename_map[col] = "TARIFF"
        elif u in ["MOD", "MODIFIER"]:
            rename_map[col] = "MODIFIER"
        elif u in ["QTY", "QUANTITY"]:
            rename_map[col] = "QTY"
        elif u in ["FEES", "FEE"]:
            rename_map[col] = "FEES"
        elif u in ["AMOUNT", "AMT", "TOTAL", "COST"]:
            rename_map[col] = "AMOUNT"

    df = df.rename(columns=rename_map)

    # If FEES missing and AMOUNT exists, use AMOUNT as FEES (UCS.xlsx pattern)
    if "FEES" not in df.columns and "AMOUNT" in df.columns:
        df["FEES"] = df["AMOUNT"]

    # Ensure required columns exist (create empty if missing)
    required = ["DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"]
    for r in required:
        if r not in df.columns:
            df[r] = ""

    # Clean numeric fields
    def to_number(x):
        if pd.isna(x) or x == "":
            return 0.0
        s = str(x).replace(",", "").replace("$", "").strip()
        try:
            return float(s)
        except:
            # sometimes the column holds non-numeric tokens; return 0
            return 0.0

    def to_qty(x):
        if pd.isna(x) or x == "":
            return 1
        s = str(x).strip()
        # remove .0 if present
        if s.replace(".", "", 1).isdigit():
            try:
                return int(float(s))
            except:
                return 1
        return 1

    df["QTY"] = df["QTY"].apply(to_qty)
    df["FEES"] = df["FEES"].apply(to_number)
    df["AMOUNT"] = df["AMOUNT"].apply(to_number)

    # Build rows, skipping TOTAL and empty descriptions
    rows = []
    for idx, r in df.iterrows():
        desc = str(r["DESCRIPTION"]).strip()
        if desc == "" or desc.upper() == "TOTAL":
            continue
        rows.append({
            "INDEX": int(idx),
            "SCAN": desc,
            "TARIFF": str(r["TARIFF"]).strip(),
            "MODIFIER": str(r["MODIFIER"]).strip(),
            "QTY": int(r["QTY"]) if pd.notna(r["QTY"]) else 1,
            "FEES": float(r["FEES"]),
            "AMOUNT": float(r["AMOUNT"])
        })

    return rows

# ------------------------------------------------------------
# Fill Excel template
# ------------------------------------------------------------
def fill_excel_template(template_file, patient, member, provider, scan_type, rows):
    """
    Writes rows to the template and returns BytesIO workbook.
    rows is a list of dicts with keys:
    SCAN, TARIFF, MODIFIER, QTY, FEES, AMOUNT
    """
    workbook = openpyxl.load_workbook(template_file)
    ws = workbook.active

    # Column mapping (A=1 ... G=7)
    colmap = {
        "DESCRIPTION": 1,  # A
        "TARIFF": 2,       # B
        "MODIFIER": 3,     # C
        "QTY": 5,          # E
        "FEES": 6,         # F
        "AMOUNT": 7        # G
    }

    # Header info (same rows you used before)
    write_force(ws, 13, 2, f"FOR PATIENT: {patient}")
    write_force(ws, 14, 2, f"MEMBER NUMBER: {member}")
    write_force(ws, 12, 5, provider)
    # Also write scan type somewhere visible (optional) - row 12 col 2
    write_force(ws, 12, 2, f"SCAN TYPE: {scan_type}")

    # Write rows starting at row 20
    rowptr = 20
    for r in rows:
        write_force(ws, rowptr, colmap["DESCRIPTION"], r["SCAN"])
        write_force(ws, rowptr, colmap["TARIFF"], r["TARIFF"])
        write_force(ws, rowptr, colmap["MODIFIER"], r["MODIFIER"])
        write_force(ws, rowptr, colmap["QTY"], r["QTY"])
        write_force(ws, rowptr, colmap["FEES"], r["FEES"])
        write_force(ws, rowptr, colmap["AMOUNT"], r["AMOUNT"])
        rowptr += 1

    # Total to G22 (row 22, col 7)
    total = sum([float(x["AMOUNT"]) for x in rows])
    write_force(ws, 22, 7, total)

    out = io.BytesIO()
    workbook.save(out)
    out.seek(0)
    return out

# ------------------------------------------------------------
# Streamlit UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

template = st.file_uploader("Upload Excel Template", type=["xlsx"])
charge = st.file_uploader("Upload Charge Sheet (UCS style)", type=["xlsx"])
patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Provider / ATT")

# Scan type selector (fixed list)
scan_type = st.selectbox("Select Scan Type for this Quotation", ["-- choose --"] + SCAN_TYPES)

# Session storage
if "parsed_rows" not in st.session_state:
    st.session_state.parsed_rows = None

# Parse button
col1, col2 = st.columns([1, 1])
with col1:
    if charge and st.button("Parse Charge Sheet"):
        try:
            rows = parse_charge_sheet(charge)
            st.session_state.parsed_rows = rows
            st.success(f"Parsed {len(rows)} rows from charge sheet.")
        except Exception as e:
            st.session_state.parsed_rows = None
            st.error(f"Error parsing charge sheet: {e}")

# If parsed, show preview and multi-select for rows to include
if st.session_state.parsed_rows:
    st.subheader("Parsed Charge Sheet Preview")
    df_preview = pd.DataFrame(st.session_state.parsed_rows)
    # show useful columns
    display_df = df_preview[["INDEX", "SCAN", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"]].copy()
    display_df = display_df.reset_index(drop=True)
    st.dataframe(display_df, use_container_width=True)

    # Multi-select to pick rows to include (label uses original dataframe INDEX for robustness)
    options = [
        f"{r['INDEX']} — {r['SCAN']} — {r['TARIFF']}"
        for r in st.session_state.parsed_rows
    ]
    default_opts = options.copy()  # default: select all
    chosen = st.multiselect("Select rows to include in quotation", options, default=default_opts)

    # Compute selected rows
    selected_rows = []
    selected_indices = set()
    for sel in chosen:
        idx_part = sel.split("—")[0].strip()
        try:
            idx = int(idx_part)
            selected_indices.add(idx)
        except:
            continue

    for r in st.session_state.parsed_rows:
        if r["INDEX"] in selected_indices:
            selected_rows.append({
                "SCAN": r["SCAN"],
                "TARIFF": r["TARIFF"],
                "MODIFIER": r["MODIFIER"],
                "QTY": r["QTY"],
                "FEES": r["FEES"],
                "AMOUNT": r["AMOUNT"]
            })

    st.markdown(f"**Rows selected:** {len(selected_rows)}")
    if len(selected_rows) > 0:
        st.dataframe(pd.DataFrame(selected_rows), use_container_width=True)

    # Generate button (right column)
    with col2:
        if template and patient and member and provider and scan_type != "-- choose --":
            if st.button("Generate Quotation"):
                try:
                    output = fill_excel_template(template, patient, member, provider, scan_type, selected_rows)
                    st.success("Quotation generated successfully.")
                    st.download_button(
                        "Download Quotation",
                        output,
                        file_name=f"quotation_{patient.replace(' ','_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Error generating quotation: {e}")
        else:
            st.info("Upload Template, fill Patient/Member/Provider, choose Scan Type, and select rows to enable Generate.")

else:
    st.info("Upload a charge sheet and press 'Parse Charge Sheet' to preview rows.")
