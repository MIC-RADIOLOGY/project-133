import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------------------------
# Fixed scan-type list
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
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            tl = mr.coord.split(":")[0]
            cell = ws[tl]
            break
    cell.value = value

# ------------------------------------------------------------
# Parse charge sheet (robust + guaranteed columns)
# ------------------------------------------------------------
def parse_charge_sheet(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)

    aliases = {
        "DESCRIPTION": ["DESCRIPTION", "EXAMINATION", "SCAN", "ITEM"],
        "TARIFF": ["TARRIF", "TARIFF"],
        "MODIFIER": ["MODIFIER", "MOD"],
        "QTY": ["QTY", "QUANTITY"],
        "FEES": ["FEES", "FEE"],
        "AMOUNT": ["AMOUNT", "AMT", "TOTAL", "COST"]
    }

    header_row = None
    for i in range(len(df_raw)):
        row = [str(x).strip().upper() if pd.notna(x) else "" for x in df_raw.iloc[i].tolist()]
        found = sum(1 for name_list in aliases.values() for n in name_list if n in row)
        if found >= 3:
            header_row = i
            break

    if header_row is None:
        raise ValueError(
            "Could not locate a header row. Ensure the charge sheet contains a header like EXAMINATION / TARRIF / MODIFIER / QUANTITY / AMOUNT"
        )

    df = pd.read_excel(uploaded_file, header=header_row, dtype=str)
    df.columns = [str(c).strip().upper() for c in df.columns]

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

    # Ensure all required columns exist
    required = ["DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"]
    for r in required:
        if r not in df.columns:
            df[r] = ""

    # Fill FEES if missing
    if df["FEES"].eq("").all() and "AMOUNT" in df.columns:
        df["FEES"] = df["AMOUNT"]

    # Numeric conversion
    def to_number(x):
        if pd.isna(x) or x == "":
            return 0.0
        s = str(x).replace(",", "").replace("$", "").strip()
        try:
            return float(s)
        except:
            return 0.0

    def to_qty(x):
        if pd.isna(x) or x == "":
            return 1
        s = str(x).strip()
        if s.replace(".", "", 1).isdigit():
            try:
                return int(float(s))
            except:
                return 1
        return 1

    df["QTY"] = df["QTY"].apply(to_qty)
    df["FEES"] = df["FEES"].apply(to_number)
    df["AMOUNT"] = df["AMOUNT"].apply(to_number)

    # Build rows
    rows = []
    for idx, r in df.iterrows():
        desc = str(r["DESCRIPTION"]).strip()
        if desc == "" or desc.upper() == "TOTAL":
            continue
        rows.append({
            "INDEX": int(idx),
            "SCAN": desc,
            "TARIFF": str(r.get("TARIFF", "")).strip(),
            "MODIFIER": str(r.get("MODIFIER", "")).strip(),
            "QTY": int(r.get("QTY", 1)),
            "FEES": float(r.get("FEES", 0.0)),
            "AMOUNT": float(r.get("AMOUNT", 0.0))
        })

    # Ensure at least one row exists
    if len(rows) == 0:
        rows.append({
            "INDEX": 0,
            "SCAN": "",
            "TARIFF": "",
            "MODIFIER": "",
            "QTY": 1,
            "FEES": 0.0,
            "AMOUNT": 0.0
        })

    return rows

# ------------------------------------------------------------
# Fill Excel template
# ------------------------------------------------------------
def fill_excel_template(template_file, patient, member, provider, scan_type, rows):
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

    write_force(ws, 13, 2, f"FOR PATIENT: {patient}")
    write_force(ws, 14, 2, f"MEMBER NUMBER: {member}")
    write_force(ws, 12, 5, provider)
    write_force(ws, 12, 2, f"SCAN TYPE: {scan_type}")

    rowptr = 20
    for r in rows:
        write_force(ws, rowptr, colmap["DESCRIPTION"], r["SCAN"])
        write_force(ws, rowptr, colmap["TARIFF"], r["TARIFF"])
        write_force(ws, rowptr, colmap["MODIFIER"], r["MODIFIER"])
        write_force(ws, rowptr, colmap["QTY"], r["QTY"])
        write_force(ws, rowptr, colmap["FEES"], r["FEES"])
        write_force(ws, rowptr, colmap["AMOUNT"], r["AMOUNT"])
        rowptr += 1

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

scan_type = st.selectbox("Select Scan Type for this Quotation", ["-- choose --"] + SCAN_TYPES)

if "parsed_rows" not in st.session_state:
    st.session_state.parsed_rows = None

col1, col2 = st.columns([1, 1])

with col1:
    if charge and st.button("Parse Charge Sheet"):
        try:
            rows = parse_charge_sheet(charge)
            # Sanitize rows: ensure all required keys exist
            for i, r in enumerate(rows):
                for k in ["INDEX", "SCAN", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"]:
                    if k not in r:
                        if k == "INDEX":
                            r[k] = i
                        elif k == "SCAN" or k == "TARIFF" or k == "MODIFIER":
                            r[k] = ""
                        elif k == "QTY":
                            r[k] = 1
                        else:
                            r[k] = 0.0
            st.session_state.parsed_rows = rows
            st.success(f"Parsed {len(rows)} rows from charge sheet.")
        except Exception as e:
            st.session_state.parsed_rows = None
            st.error(f"Error parsing charge sheet: {e}")

if st.session_state.parsed_rows:
    st.subheader("Parsed Charge Sheet Preview")
    df_preview = pd.DataFrame(st.session_state.parsed_rows)
    display_columns = ["INDEX", "SCAN", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT"]
    display_columns = [c for c in display_columns if c in df_preview.columns]
    display_df = df_preview[display_columns].copy().reset_index(drop=True)
    st.dataframe(display_df, use_container_width=True)

    options = []
    for r in st.session_state.parsed_rows:
        index = r.get("INDEX", 0)
        scan = r.get("SCAN", "")
        tariff = r.get("TARIFF", "")
        options.append(f"{index} — {scan} — {tariff}")

    default_opts = options.copy()
    chosen = st.multiselect("Select rows to include in quotation", options, default=default_opts)

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
        if r.get("INDEX", -1) in selected_indices:
            selected_rows.append({
                "SCAN": r.get("SCAN", ""),
                "TARIFF": r.get("TARIFF", ""),
                "MODIFIER": r.get("MODIFIER", ""),
                "QTY": r.get("QTY", 1),
                "FEES": r.get("FEES", 0.0),
                "AMOUNT": r.get("AMOUNT", 0.0)
            })

    st.markdown(f"**Rows selected:** {len(selected_rows)}")
    if len(selected_rows) > 0:
        st.dataframe(pd.DataFrame(selected_rows), use_container_width=True)

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
