# app.py
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
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")

    if st.button("Login"):
        if u == "admin" and p == "Jamela2003":
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
# PARSER (FIXED COLUMN SHIFT)
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

        exam_u = exam.upper()

        if exam_u.endswith("SCAN") or exam_u in {"XRAY", "MRI", "ULTRASOUND"}:
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

        raw_qty = clean_text(r["D_QTY"])
        raw_amount = clean_text(r["E_AMOUNT"])

        qty = safe_int(raw_qty, 1)
        amount = safe_float(raw_amount, 0.0)

        # ---- FIX MISALIGNED FEES ----
        # If qty is unrealistically large, it is actually FEES
        if qty > 20 and amount == 0:
            fees = qty
            qty = 1
            amount = fees
        else:
            fees = amount / qty if qty else amount

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "IS_MAIN_SCAN": exam_u not in COMPONENT_KEYS,
            "TARIFF": safe_float(r["B_TARIFF"], None),
            "MODIFIER": clean_text(r["C_MOD"]),
            "QTY": qty,
            "FEES": round(fees, 2),
            "AMOUNT": round(amount, 2)
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
                ws.cell(row=mr.min_row, column=mr.min_col).value = value
                return

def find_template_positions(ws):
    pos = {}
    headers = ["DESCRIPTION", "TARIFF", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]

    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if not cell.value:
                continue
            t = str(cell.value).upper().strip()

            if any(h in t for h in headers):
                pos.setdefault("cols", {})
                pos.setdefault("table_start_row", cell.row + 1)
                for h in headers:
                    if h in t:
                        pos["cols"].setdefault(h, []).append(cell.column)
    return pos

# ------------------------------------------------------------
# TEMPLATE FILL
# ------------------------------------------------------------
def fill_excel_template(template_file, rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    rowptr = pos.get("table_start_row", 22)
    total = 0.0

    for r in rows:
        desc = r["SCAN"]
        if not r["IS_MAIN_SCAN"]:
            desc = "   " + desc

        for c in pos["cols"].get("DESCRIPTION", []):
            write_safe(ws, rowptr, c, desc)

        for c in pos["cols"].get("TARIFF", []) + pos["cols"].get("TARRIF", []):
            write_safe(ws, rowptr, c, r["TARIFF"])

        for c in pos["cols"].get("MOD", []):
            write_safe(ws, rowptr, c, r["MODIFIER"])

        for c in pos["cols"].get("QTY", []):
            write_safe(ws, rowptr, c, r["QTY"])

        for c in pos["cols"].get("FEES", []):
            write_safe(ws, rowptr, c, r["FEES"])

        for c in pos["cols"].get("AMOUNT", []):
            write_safe(ws, rowptr, c, r["AMOUNT"])

        total += r["AMOUNT"]
        rowptr += 1

    for c in pos["cols"].get("AMOUNT", []):
        write_safe(ws, 22, c, round(total, 2))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

charge = st.file_uploader("Upload Charge Sheet", type=["xlsx"])
template = st.file_uploader("Upload Template", type=["xlsx"])

if charge and st.button("Load Sheet"):
    st.session_state.df = load_charge_sheet(charge)

if "df" in st.session_state:
    df = st.session_state.df

    selected = st.multiselect(
        "Select scans",
        df.index,
        format_func=lambda i: f"{df.at[i,'SCAN']} | {df.at[i,'AMOUNT']}"
    )

    if selected:
        editor_df = df.loc[selected, [
            "SCAN", "TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT", "IS_MAIN_SCAN"
        ]].reset_index(drop=True)

        st.subheader("Edit Descriptions")
        edited = st.data_editor(
            editor_df,
            disabled=["TARIFF", "MODIFIER", "QTY", "FEES", "AMOUNT", "IS_MAIN_SCAN"],
            use_container_width=True
        )

        st.subheader("Preview (Read-Only)")
        st.dataframe(edited)

        if template and st.button("Generate Excel"):
            out = fill_excel_template(template, edited.to_dict("records"))
            st.download_button(
                "Download Quotation",
                out,
                "quotation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
