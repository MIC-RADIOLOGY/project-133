import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ------------------------------------------------------------
# LOGIN
# ------------------------------------------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("Login Required")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username == "admin" and password == "Jamela2003":
            st.session_state.logged_in = True
            st.success("Login successful")
        else:
            st.error("Invalid credentials")
    st.stop()

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
COMPONENT_KEYS = {
    "PELVIS", "CONSUMABLES", "FF",
    "IV", "IV CONTRAST", "IV CONTRAST 100MLS"
}

GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

MAIN_CATEGORIES = set()

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def safe_int(x, default=1):
    try:
        return int(float(str(x).replace(",", "").strip()))
    except Exception:
        return default

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "").strip())
    except Exception:
        return default

# ------------------------------------------------------------
# PARSER
# ------------------------------------------------------------
def load_charge_sheet(file):
    df_raw = pd.read_excel(file, header=None, dtype=object)

    while df_raw.shape[1] < 5:
        df_raw[df_raw.shape[1]] = None

    df_raw = df_raw.iloc[:, :5]
    df_raw.columns = ["A_EXAM", "B_TARIFF", "C_MOD", "D_QTY", "E_AMOUNT"]

    structured = []
    current_category = None
    current_subcategory = None

    for _, r in df_raw.iterrows():
        exam = clean_text(r["A_EXAM"])
        if not exam:
            continue

        exam_u = exam.upper().strip()

        if exam_u in MAIN_CATEGORIES or exam_u.endswith("SCAN") or exam_u in {"XRAY", "MRI", "ULTRASOUND"}:
            MAIN_CATEGORIES.add(exam_u)
            current_category = exam
            current_subcategory = None
            continue

        if exam_u in GARBAGE_KEYS:
            continue

        if clean_text(r["B_TARIFF"]) == "" and clean_text(r["E_AMOUNT"]) == "":
            current_subcategory = exam
            continue

        if not current_category:
            continue

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "IS_MAIN_SCAN": exam_u not in COMPONENT_KEYS,
            "TARIFF": safe_float(r["B_TARIFF"]),
            "MODIFIER": clean_text(r["C_MOD"]),
            "QTY": safe_int(r["D_QTY"]),
            "AMOUNT": safe_float(r["E_AMOUNT"])
        })

    return pd.DataFrame(structured)

# ------------------------------------------------------------
# EXCEL POSITION SCAN (TEMPLATE-AWARE)
# ------------------------------------------------------------
def find_template_positions(ws):
    pos = {
        "COLS": {},
        "MOD_COL": None
    }

    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if not cell.value:
                continue

            t = str(cell.value).upper()

            if "DESCRIPTION" in t:
                pos["COLS"]["DESC"] = cell.column
                pos["START_ROW"] = cell.row + 1

            if "TARIFF" in t or "TARRIF" in t:
                pos["COLS"]["TARIFF"] = cell.column

            if "QTY" in t:
                pos["COLS"]["QTY"] = cell.column

            if "FEE" in t:
                pos["COLS"]["FEES"] = cell.column

            if "AMOUNT" in t:
                pos["COLS"]["AMOUNT"] = cell.column

            # 🔒 CRITICAL FIX: FIRST MOD ONLY
            if "MOD" in t and pos["MOD_COL"] is None:
                pos["MOD_COL"] = cell.column

            if "PATIENT" in t:
                pos["PATIENT"] = (cell.row, cell.column)

            if "MEMBER" in t:
                pos["MEMBER"] = (cell.row, cell.column)

            if "PROVIDER" in t or "MEDICAL AID" in t:
                pos["PROVIDER"] = (cell.row, cell.column)

            if t.strip() == "DATE":
                pos["DATE"] = (cell.row, cell.column)

    return pos

# ------------------------------------------------------------
# SAFE WRITE (MERGED-CELL SAFE)
# ------------------------------------------------------------
def write_safe(ws, r, c, value):
    if not c:
        return
    try:
        ws.cell(row=r, column=c).value = value
    except Exception:
        for mr in ws.merged_cells.ranges:
            if ws.cell(r, c).coordinate in mr:
                ws.cell(mr.min_row, mr.min_col).value = value
                return

# ------------------------------------------------------------
# TEMPLATE FILL
# ------------------------------------------------------------
def fill_excel_template(template_file, patient, member, provider, scan_rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    if "PATIENT" in pos:
        write_safe(ws, *pos["PATIENT"], patient)

    if "MEMBER" in pos:
        write_safe(ws, *pos["MEMBER"], member)

    if "PROVIDER" in pos:
        write_safe(ws, *pos["PROVIDER"], provider)

    if "DATE" in pos:
        write_safe(ws, pos["DATE"][0] + 1, pos["DATE"][1],
                   datetime.today().strftime("%d/%m/%Y"))

    rowptr = pos.get("START_ROW", 22)
    grand_total = 0.0

    for sr in scan_rows:
        desc = sr["SCAN"] if sr["IS_MAIN_SCAN"] else "   " + sr["SCAN"]

        write_safe(ws, rowptr, pos["COLS"]["DESC"], desc)
        write_safe(ws, rowptr, pos["COLS"].get("TARIFF"), sr["TARIFF"])
        write_safe(ws, rowptr, pos["MOD_COL"], sr["MODIFIER"])
        write_safe(ws, rowptr, pos["COLS"].get("QTY"), sr["QTY"])

        fee = sr["AMOUNT"] / sr["QTY"] if sr["QTY"] else sr["AMOUNT"]
        write_safe(ws, rowptr, pos["COLS"].get("FEES"), round(fee, 2))

        grand_total += sr["AMOUNT"]
        rowptr += 1

    write_safe(ws, 22, pos["COLS"].get("AMOUNT"), round(grand_total, 2))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ------------------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if uploaded_charge and st.button("Load & Parse Charge Sheet"):
    st.session_state.df = load_charge_sheet(uploaded_charge)
    st.success("Charge sheet parsed successfully")

if "df" in st.session_state:
    df = st.session_state.df

    # CATEGORY -> SUBCATEGORY -> SCAN checkboxes
    selected_rows = []
    st.subheader("Select scans to include in quotation")
    for main_cat in sorted(df["CATEGORY"].unique()):
        with st.expander(main_cat, expanded=False):
            subcats = df[df["CATEGORY"] == main_cat]["SUBCATEGORY"].dropna().unique()
            for subcat in subcats:
                with st.expander(f"Subcategory: {subcat}", expanded=False):
                    scans = df[(df["CATEGORY"] == main_cat) & (df["SUBCATEGORY"] == subcat)].reset_index(drop=True)
                    for idx, r in scans.iterrows():
                        label = f"{r['SCAN']} | Tariff {r['TARIFF']} | Amount {r['AMOUNT']}"
                        if st.checkbox(label, key=f"{main_cat}_{subcat}_{idx}"):
                            selected_rows.append(r.to_dict())

            # Also allow scans without subcategory
            scans_no_sub = df[(df["CATEGORY"] == main_cat) & (df["SUBCATEGORY"].isna())].reset_index(drop=True)
            for idx, r in scans_no_sub.iterrows():
                label = f"{r['SCAN']} | Tariff {r['TARIFF']} | Amount {r['AMOUNT']}"
                if st.checkbox(label, key=f"{main_cat}_none_{idx}"):
                    selected_rows.append(r.to_dict())

    if selected_rows:
        st.subheader("Preview of selected scans")
        st.dataframe(pd.DataFrame(selected_rows)[
            ["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]
        ])

        if uploaded_template and st.button("Generate & Download Quotation"):
            out = fill_excel_template(
                uploaded_template, patient, member, provider, selected_rows
            )

            st.download_button(
                "Download Quotation",
                data=out,
                file_name="quotation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
