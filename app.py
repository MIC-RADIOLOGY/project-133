# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
from typing import Optional

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN", "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND"
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
        x_str = str(x).replace(",", "").strip()
        return int(float(x_str))
    except Exception:
        return default

def safe_float(x, default=0.0):
    try:
        x_str = str(x).replace(",", "").strip()
        return float(x_str)
    except Exception:
        return default

# ---------- Parser ----------
def load_charge_sheet(file) -> pd.DataFrame:
    df_raw = pd.read_excel(file, header=None, dtype=object)
    while df_raw.shape[1] < 5:
        df_raw[df_raw.shape[1]] = None
    df_raw = df_raw.iloc[:, :5]
    df_raw.columns = ["A_EXAM", "B_TARIFF", "C_MOD", "D_QTY", "E_AMOUNT"]

    structured = []
    current_category: Optional[str] = None
    current_subcategory: Optional[str] = None

    for _, r in df_raw.iterrows():
        exam = clean_text(r["A_EXAM"])
        exam_u = exam.upper()

        # Skip rows where DESCRIPTION is empty and not FF
        if exam == "" and exam_u != "FF":
            continue

        # MAIN CATEGORY
        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        # Special case: FF
        if exam_u == "FF":
            row_tariff = safe_float(r["B_TARIFF"], default=None)
            row_amt = safe_float(r["E_AMOUNT"], default=0.0)
            row_qty = safe_int(r["D_QTY"], default=1)
            structured.append({
                "CATEGORY": current_category,
                "SUBCATEGORY": current_subcategory,
                "SCAN": "FF",
                "TARIFF": row_tariff,
                "MODIFIER": "",
                "QTY": row_qty,
                "AMOUNT": row_amt
            })
            continue

        # Skip garbage keys
        if exam_u in GARBAGE_KEYS:
            continue

        # Subcategory row (tariff & amount blank)
        tariff_str = str(r["B_TARIFF"]).strip() if not pd.isna(r["B_TARIFF"]) else ""
        amount_str = str(r["E_AMOUNT"]).strip() if not pd.isna(r["E_AMOUNT"]) else ""
        if tariff_str in ["", "nan", "None", "NaN"] and amount_str in ["", "nan", "None", "NaN"]:
            current_subcategory = exam
            continue

        # Only rows with DESCRIPTION text reach this point
        row_tariff = safe_float(r["B_TARIFF"], default=None)
        row_amt = safe_float(r["E_AMOUNT"], default=0.0)
        row_qty = safe_int(r["D_QTY"], default=1)
        row_mod = clean_text(r["C_MOD"])

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "TARIFF": row_tariff,
            "MODIFIER": row_mod,
            "QTY": row_qty,
            "AMOUNT": row_amt
        })

    return pd.DataFrame(structured)

# ---------- Excel template filler ----------
def write_safe(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    try:
        cell.value = value
    except Exception:
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                top = mr.coord.split(":")[0]
                ws[top].value = value
                return

def find_template_positions(ws):
    pos = {}
    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if cell.value:
                t = u(cell.value)
                if "PATIENT" in t and "patient_cell" not in pos:
                    pos["patient_cell"] = (cell.row, cell.column + 1)
                if "MEMBER" in t and "member_cell" not in pos:
                    pos["member_cell"] = (cell.row, cell.column + 1)
                if ("PROVIDER" in t or "EXAMINATION" in t) and "provider_cell" not in pos:
                    pos["provider_cell"] = (cell.row, cell.column + 1)

                headers = ["DESCRIPTION","TARRIF","MOD","QTY","FEES","AMOUNT"]
                if any(h in t for h in headers) and "cols" not in pos:
                    pos["cols"] = {}
                    pos["table_start_row"] = cell.row + 1
                for h in headers:
                    if h in t:
                        pos["cols"][h] = cell.column
    return pos

def fill_excel_template(template_file, patient, member, provider, scan_rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    if "patient_cell" in pos:
        r, c = pos["patient_cell"_]()
