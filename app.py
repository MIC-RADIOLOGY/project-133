# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN",
    "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND"
}

GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

COMPONENT_KEYS = {
    "PELVIS", "CONSUMABLES", "FF",
    "IV", "IV CONTRAST", "IV CONTRAST 100MLS"
}

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

        # Main category
        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        # Skip totals / garbage
        if exam_u in GARBAGE_KEYS:
            continue

        # Subcategory rows (no tariff & no amount)
        if clean_text(r["B_TARIFF"]) == "" and clean_text(r["E_AMOUNT"]) == "":
            current_subcategory = exam
            continue

        if not current_category:
            continue

        # ✅ Correct main/component scan logic
        is_main_scan = exam_u not in COMPONENT_KEYS

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,                      # EXACT text preserved
            "IS_MAIN_SCAN": is_main_scan,
            "TARIFF": safe_float(r["B_TARIFF"], None),
            "MODIFIER": str(clean_text(r["C_MOD"])),
            "QTY": safe_int(r["D_QTY"], 1),
            "AMOUNT": safe_float(r["E_AMOUNT"], 0.0)
        })

    return pd.DataFrame(structured)

# ------------------------------------------------------------
# EXCEL HELPERS
# ------------------------------------------------------------
def write_safe(ws, r, c, value):
    if not c:
        return
    cell = ws.cell(row=r, column=c)
    try:
        cell.value = value
    except Exception:
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                start_cell = ws.cell(row=mr.min_row, column=mr.min_col)
                start_cell.value = value
                return

def append_after_label(ws, r, c, label, value):
    if not value:
        return
    cell = ws.cell(row=r, column=c)
    existing = str(cell.value) if cell.value else ""
    cell.value = f"{existing.strip()} {value}".strip()

def write_below_label(ws, r, c, value):
    target = ws.cell(row=r + 1, column=c)
    try:
        target.value = value
    except Exception:
        for mr in ws.merged_cells.ranges:
            if target.coordinate in mr:
                start_cell = ws.cell(row=mr.min_row, column=mr.min_col)
                start_cell.value = value
                return

def find_template_positions(ws):
    pos = {}
    headers = ["DESCRIPTION", "TARIFF", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]

    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if not cell.value:
                continue

            t = str(cell.value).upper()

            if "PATIENT" in t:
                pos.setdefault("patient_cell", (cell.row, cell.column))
            if "MEMBER" in t:
                pos.setdefault("member_cell", (cell.row, cell.column))
            if "PROVIDER" in t or "MEDICAL AID" in t:
                pos.setdefault("provider_cell", (cell.row, cell.column))
            if t.strip() == "DATE":
                pos.setdefault("date_cell", (cell.row, cell.column))

            if any(h in t for h in headers):
                pos.setdefault("cols", {})
                pos.setdefault("table_start_row", cell.row + 1)
                for h in headers:
                    if h in t:
                        pos["cols"][h] = cell.column
    return pos

# ------------------------------------------------------------
# TEMPLATE FILL
# ------------------------------------------------------------
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

    if "date_cell" in pos:
        write_below_label(ws, *pos["date_cell"],
                          datetime.today().strftime("%d/%m/%Y"))

    rowptr = pos.get("table_start_row", 22)
    grand_total = 0.0

    for sr in scan_rows:
        if sr["IS_MAIN_SCAN"]:
            write_safe(ws, rowptr, pos["cols"].get("DESCRIPTION"), sr["SCAN"])

        write_safe(ws, rowptr,
                   pos["cols"].get("TARIFF") or pos["cols"].get("TARRIF"),
                   sr["TARIFF"])

        write_safe(ws, rowptr, pos["cols"].get("MOD"), sr["MODIFIER"])
        write_safe(ws, rowptr, pos["cols"].get("QTY"), sr["QTY"])

        fees = sr["AMOUNT"] / sr["QTY"] if sr["QTY"] else sr["AMOUNT"]
        write_safe(ws, rowptr, pos["cols"].get("FEES"), round(fees, 2))

        grand_total += sr["AMOUNT"]
        rowptr += 1

    # Only one total — CELL G22
    write_safe(ws, 22, pos["cols"].get("AMOUNT"), round(grand_total, 2))

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
    st.success("Charge sheet parsed successfully.")

if "df" in st.session_state:
    df = st.session_state.df

    main_sel = st.selectbox(
        "Select Main Category",
        sorted(df["CATEGORY"].dropna().unique())
    )

    subcats = sorted(df[df["CATEGORY"] == main_sel]["SUBCATEGORY"].dropna().unique())
    sub_sel = st.selectbox("Select Subcategory", subcats) if subcats else None

    scans = (
        df[(df["CATEGORY"] == main_sel) & (df["SUBCATEGORY"] == sub_sel)]
        if sub_sel else df[df["CATEGORY"] == main_sel]
    ).reset_index(drop=True)

    scans["label"] = scans.apply(
        lambda r: f"{r['SCAN']} | Tariff {r['TARIFF']} | Amount {r['AMOUNT']}",
        axis=1
    )

    selected = st.multiselect(
        "Select scans to include",
        options=list(range(len(scans))),
        format_func=lambda i: scans.at[i, "label"]
    )

    selected_rows = [scans.iloc[i].to_dict() for i in selected]

    if selected_rows:
        st.dataframe(pd.DataFrame(selected_rows)[
            ["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]
        ])

        if uploaded_template and st.button("Generate & Download Quotation"):
            safe_name = "".join(
                c for c in (patient or "patient")
                if c.isalnum() or c in (" ", "_")
            ).strip()

            out = fill_excel_template(
                uploaded_template, patient, member, provider, selected_rows
            )

            st.download_button(
                "Download Quotation",
                data=out,
                file_name=f"quotation_{safe_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
