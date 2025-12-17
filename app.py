# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
from typing import Optional

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

        if any(k in exam_u for k in MAIN_CATEGORIES):
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
        amount_blank = pd.isna(r["E_AMOUNT"]) or str(r["E_AMOUNT"]).strip() == ""

        if tariff_blank and amount_blank:
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
                ws.cell(mr.min_row, mr.min_col).value = value
                return

def append_to_cell(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    text = str(cell.value or "")
    if ":" in text:
        label = text.split(":", 1)[0].strip()
        cell.value = f"{label}: {value}"
    else:
        cell.value = f"{text} {value}"

# ---------- Template scanning ----------
def find_template_positions(ws):
    pos = {}
    headers = ["DESCRIPTION", "TARIFF", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]

    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if not cell.value:
                continue

            t = u(cell.value)

            if "FOR PATIENT" in t and "patient_cell" not in pos:
                pos["patient_cell"] = (cell.row, cell.column)

            if "MEMBER NUMBER" in t and "member_cell" not in pos:
                pos["member_cell"] = (cell.row, cell.column)

            if "ATT" in t and "att_cell" not in pos:
                pos["att_cell"] = (cell.row, cell.column)

            if any(h in t for h in headers):
                pos.setdefault("cols", {})
                pos.setdefault("table_start_row", cell.row + 1)

                for h in headers:
                    if h in t:
                        key = "TARIFF" if h == "TARRIF" else h
                        pos["cols"][key] = cell.column

    return pos

# ---------- Fill template ----------
def fill_excel_template(template_file, patient, member, provider, scan_rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    if "patient_cell" in pos:
        append_to_cell(ws, *pos["patient_cell"], patient)

    if "member_cell" in pos:
        append_to_cell(ws, *pos["member_cell"], member)

    if "att_cell" in pos:
        append_to_cell(ws, *pos["att_cell"], provider)

    if "table_start_row" in pos and "cols" in pos:
        r = pos["table_start_row"]
        cols = pos["cols"]

        for sr in scan_rows:
            write_safe(ws, r, cols.get("DESCRIPTION"), sr["SCAN"])
            write_safe(ws, r, cols.get("TARIFF"), sr["TARIFF"])
            write_safe(ws, r, cols.get("MOD"), sr["MODIFIER"])
            write_safe(ws, r, cols.get("QTY"), sr["QTY"])
            write_safe(ws, r, cols.get("FEES"), sr["TARIFF"] * sr["QTY"])
            write_safe(ws, r, cols.get("AMOUNT"), sr["AMOUNT"])
            r += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------- Streamlit UI ----------
st.title("📄 Medical Quotation Generator")

debug = st.checkbox("Show debug output")

charge_file = st.file_uploader("Upload Charge Sheet", type="xlsx")
template_file = st.file_uploader("Upload Quotation Template", type="xlsx")

patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if charge_file and st.button("Load & Parse Charge Sheet"):
    st.session_state.df = load_charge_sheet(charge_file)
    st.success("Charge sheet parsed successfully")

if "df" in st.session_state:
    df = st.session_state.df

    if debug:
        st.dataframe(df)

    cats = sorted(df["CATEGORY"].dropna().unique())
    cat = st.selectbox("Select Main Category", cats)

    subs = sorted(df[df["CATEGORY"] == cat]["SUBCATEGORY"].dropna().unique())
    sub = st.selectbox("Select Subcategory", subs) if subs else None

    data = df[(df["CATEGORY"] == cat) & (df["SUBCATEGORY"] == sub)] if sub else df[df["CATEGORY"] == cat]

    data = data.reset_index(drop=True)
    data["label"] = data.apply(
        lambda r: f"{r['SCAN']} | Tariff {r['TARIFF']} | Amount {r['AMOUNT']}", axis=1
    )

    sel = st.multiselect(
        "Select scans",
        options=list(range(len(data))),
        format_func=lambda i: data.loc[i, "label"]
    )

    selected = [data.loc[i].to_dict() for i in sel]

    if selected:
        st.dataframe(pd.DataFrame(selected)[["SCAN", "TARIFF", "QTY", "AMOUNT"]])
        total = sum(s["AMOUNT"] for s in selected)
        st.markdown(f"### Total: **{total:.2f}**")

        if template_file and st.button("Generate Quotation"):
            output = fill_excel_template(template_file, patient, member, provider, selected)
            st.download_button(
                "Download Excel",
                output,
                file_name=f"quotation_{patient or 'patient'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
