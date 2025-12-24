import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config / heuristics ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN",
    "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND"
}

GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# ---------- Helpers ----------
def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def u(x):
    return clean_text(x).upper()

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

# ---------- Parser ----------
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

        exam_u = exam.upper()

        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        if exam_u == "FF":
            structured.append({
                "CATEGORY": current_category,
                "SUBCATEGORY": current_subcategory,
                "SCAN": "FF",
                "TARIFF": safe_float(r["B_TARIFF"], None),
                "MODIFIER": "",
                "QTY": safe_int(r["D_QTY"], 1),
                "AMOUNT": safe_float(r["E_AMOUNT"], 0.0)
            })
            continue

        if exam_u in GARBAGE_KEYS:
            continue

        if clean_text(r["B_TARIFF"]) == "" and clean_text(r["E_AMOUNT"]) == "":
            current_subcategory = exam
            continue

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "TARIFF": safe_float(r["B_TARIFF"], None),
            "MODIFIER": clean_text(r["C_MOD"]),
            "QTY": safe_int(r["D_QTY"], 1),
            "AMOUNT": safe_float(r["E_AMOUNT"], 0.0)
        })

    return pd.DataFrame(structured)

# ---------- Excel helpers ----------
def write_safe(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    try:
        cell.value = value
    except Exception:
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                ws[mr.coord.split(":")[0]].value = value
                return

def append_after_label(ws, r, c, label, value):
    if not value:
        return
    cell = ws.cell(row=r, column=c)
    existing = str(cell.value) if cell.value else ""
    if label.upper() in existing.upper():
        cell.value = f"{existing.strip()} {value}"
    else:
        cell.value = value

def write_below_label(ws, r, c, value):
    if not value:
        return
    target = ws.cell(row=r + 1, column=c)
    try:
        target.value = value
    except Exception:
        for mr in ws.merged_cells.ranges:
            if target.coordinate in mr:
                ws[mr.coord.split(":")[0]].value = value
                return

def find_template_positions(ws):
    pos = {}
    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if not cell.value:
                continue

            t = u(cell.value)

            if "PATIENT" in t and "patient_cell" not in pos:
                pos["patient_cell"] = (cell.row, cell.column)

            if "MEMBER" in t and "member_cell" not in pos:
                pos["member_cell"] = (cell.row, cell.column)

            if ("PROVIDER" in t or "MEDICAL AID" in t) and "provider_cell" not in pos:
                pos["provider_cell"] = (cell.row, cell.column)

            if t.strip() == "DATE" and "date_cell" not in pos:
                pos["date_cell"] = (cell.row, cell.column)

            headers = ["DESCRIPTION", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]
            if any(h in t for h in headers):
                pos.setdefault("cols", {})
                pos.setdefault("table_start_row", cell.row + 1)
                for h in headers:
                    if h in t:
                        pos["cols"][h] = cell.column
    return pos

def fill_excel_template(template_file, patient, member, provider, scan_rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    if "patient_cell" in pos:
        append_after_label(ws, *pos["patient_cell"], "PATIENT", patient)

    if "member_cell" in pos:
        append_after_label(ws, *pos["member_cell"], "MEMBER", member)

    if "provider_cell" in pos:
        append_after_label(ws, *pos["provider_cell"], "PROVIDER", provider)

    # --- DATE below label, numeric ---
    today_numeric = datetime.today().strftime("%d/%m/%Y")
    if "date_cell" in pos:
        write_below_label(ws, *pos["date_cell"], today_numeric)

    # --- Writing the description to A22 ---
    scan_types = [sr["SCAN"] for sr in scan_rows if sr["SCAN"].lower() not in ["ff", "consumables"]]
    
    # If there are valid scan types, write the first one to A22
    if scan_types:
        write_safe(ws, 22, 1, scan_types[0])  # Cell A22 = First scan type
    
    # --- Write selected scans to the table (starting at row 23) ---
    if "table_start_row" in pos and "cols" in pos:
        rowptr = pos["table_start_row"]
        total_amount = 0.0  # To calculate total sum

        for sr in scan_rows:
            # Skip consumables and FF from being displayed in the table
            if sr["SCAN"].lower() in ["ff", "consumables"]:
                continue
            
            write_safe(ws, rowptr, pos["cols"].get("DESCRIPTION"), sr["SCAN"])
            write_safe(ws, rowptr, pos["cols"].get("TARRIF"), sr["TARIFF"])
            write_safe(ws, rowptr, pos["cols"].get("MOD"), sr["MODIFIER"])
            write_safe(ws, rowptr, pos["cols"].get("QTY"), sr["QTY"])
            write_safe(ws, rowptr, pos["cols"].get("FEES"), sr["AMOUNT"])
            total_amount += sr["AMOUNT"]
            rowptr += 1

        # --- Write total amount to F22 (Cell F22) ---
        if "cols" in pos:
            write_safe(ws, 22, pos["cols"].get("FEES"), total_amount)  # Write total to F22

    # --- Save the file and return the buffer ---
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------- Streamlit UI ----------
st.title("Medical Quotation Generator")

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if uploaded_charge and st.button("Load & Parse Charge Sheet"):
    st.session_state.df = load_charge_sheet(uploaded_charge)
    st.success("Charge sheet parsed successfully.")

if "df" in st.session_state:
    df = st.session_state.df
    categories = sorted(df["CATEGORY"].dropna().unique())

    main_sel = st.selectbox("Select Main Category", categories)
    subcats = sorted(df[df["CATEGORY"] == main_sel]["SUBCATEGORY"].dropna().unique())

    scans = (
        df[(df["CATEGORY"] == main_sel) & (df["SUBCATEGORY"] == st.selectbox("Select Subcategory", subcats))]
        if subcats else df[df["CATEGORY"] == main_sel]
    )

    scans = scans.reset_index(drop=True)
    scans["label"] = scans.apply(
        lambda r: f"{r['SCAN']} | Tariff
