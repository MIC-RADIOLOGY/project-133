import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile
from copy import copy
from openpyxl.styles import Border, Side

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config / heuristics ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN",
    "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND"
}
GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# ---------- Helpers ----------
def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def u(x) -> str:
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
def load_charge_sheet(file) -> pd.DataFrame:
    # Rewind buffer (CRITICAL on Streamlit Cloud)
    file.seek(0)

    # Validate Excel file
    try:
        zipfile.ZipFile(file)
    except zipfile.BadZipFile:
        raise ValueError("Uploaded file is not a valid .xlsx Excel file")

    file.seek(0)

    df_raw = pd.read_excel(
        file,
        header=None,
        dtype=object,
        engine="openpyxl"
    )

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

        tariff_blank = pd.isna(r["B_TARIFF"]) or str(r["B_TARIFF"]).strip() == ""
        amt_blank = pd.isna(r["E_AMOUNT"]) or str(r["E_AMOUNT"]).strip() == ""

        if tariff_blank and amt_blank:
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

# ---------- Excel Template Helpers ----------
def write_safe(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    try:
        cell.value = value
    except Exception:
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                ws[mr.coord.split(":")[0]].value = value
                return

def find_template_positions(ws):
    pos = {}
    headers = ["DESCRIPTION", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]

    for row in ws.iter_rows(min_row=1, max_row=300):
        for cell in row:
            if not cell.value:
                continue

            t = u(cell.value)

            if "PATIENT" in t and "patient_cell" not in pos:
                pos["patient_cell"] = (cell.row, cell.column)

            if "MEMBER" in t and "member_cell" not in pos:
                pos["member_cell"] = (cell.row, cell.column)

            if ("PROVIDER" in t or "EXAMINATION" in t) and "provider_cell" not in pos:
                pos["provider_cell"] = (cell.row, cell.column)

            if any(h in t for h in headers):
                pos.setdefault("cols", {})
                pos.setdefault("table_start_row", cell.row + 1)

            for h in headers:
                if h in t:
                    pos["cols"][h] = cell.column

    return pos

def replace_after_colon_in_same_cell(ws, row, col, new_value):
    cell = ws.cell(row=row, column=col)

    for rng in ws.merged_cells.ranges:
        if cell.coordinate in rng:
            cell = ws[rng.coord.split(":")[0]]
            break

    old = str(cell.value) if cell.value else ""
    cell.value = f"{old.split(':',1)[0]}: {new_value}" if ":" in old else new_value

def write_value_preserve_borders(ws, cell_address, value):
    cell = ws[cell_address]
    merged_range = None

    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            merged_range = mr
            ws.unmerge_cells(str(mr))
            cell = ws[mr.coord.split(":")[0]]
            break

    border = copy(cell.border)
    font = copy(cell.font)
    fill = copy(cell.fill)
    alignment = copy(cell.alignment)

    cell.value = value

    cell.border = border
    cell.font = font
    cell.fill = fill
    cell.alignment = alignment

    if merged_range:
        ws.merge_cells(str(merged_range))

# ---------- Fill Template ----------
def fill_excel_template(template_file, patient, member, provider, scan_rows):
    template_file.seek(0)
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    if "patient_cell" in pos:
        replace_after_colon_in_same_cell(ws, *pos["patient_cell"], patient)

    if "member_cell" in pos:
        replace_after_colon_in_same_cell(ws, *pos["member_cell"], member)

    if "provider_cell" in pos:
        replace_after_colon_in_same_cell(ws, *pos["provider_cell"], provider)

    if "table_start_row" in pos and "cols" in pos:
        rowptr = pos["table_start_row"]
        cols = pos["cols"]

        for sr in scan_rows:
            write_safe(ws, rowptr, cols.get("DESCRIPTION"), sr["SCAN"])
            write_safe(ws, rowptr, cols.get("TARRIF"), sr["TARIFF"])
            write_safe(ws, rowptr, cols.get("MOD"), sr["MODIFIER"])
            write_safe(ws, rowptr, cols.get("QTY"), sr["QTY"])
            write_safe(ws, rowptr, cols.get("FEES"), sr["AMOUNT"])
            rowptr += 1

        total_amt = sum(safe_float(r["AMOUNT"]) for r in scan_rows)
        write_value_preserve_borders(ws, "G22", total_amt)
        write_value_preserve_borders(ws, "G41", total_amt)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------- Streamlit UI ----------
st.title("Medical Quotation Generator")

debug_mode = st.checkbox("Show parsing debug output", value=False)

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if uploaded_charge and st.button("Load & Parse Charge Sheet"):
    st.session_state.parsed_df = load_charge_sheet(uploaded_charge)
    st.success("Charge sheet parsed successfully.")

if "parsed_df" in st.session_state:
    df = st.session_state.parsed_df

    if debug_mode:
        st.dataframe(df)

    cats = sorted(df["CATEGORY"].dropna().unique())
    main_sel = st.selectbox("Select Main Category", cats)
    scans = df[df["CATEGORY"] == main_sel]

    scans["label"] = scans.apply(
        lambda r: f"{r['SCAN']} | Tariff: {r['TARIFF']} | Amt: {r['AMOUNT']}",
        axis=1
    )

    sel = st.multiselect(
        "Select scans",
        scans.index,
        format_func=lambda i: scans.at[i, "label"]
    )

    selected_rows = [scans.loc[i].to_dict() for i in sel]

    if selected_rows and uploaded_template:
        out = fill_excel_template(
            uploaded_template,
            patient,
            member,
            provider,
            selected_rows
        )
        st.download_button(
            "Download Quotation",
            data=out,
            file_name=f"quotation_{patient or 'patient'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
