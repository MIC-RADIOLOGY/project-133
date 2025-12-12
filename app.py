import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN",
    "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND",
    "MRI"
}
GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# ---------- Helpers ----------
def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def safe_int(x, default=1):
    try:
        return int(float(str(x).replace(",", "").strip()))
    except:
        return default

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "").replace("$", "").strip())
    except:
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

    for idx, r in df_raw.iterrows():
        row_texts = [clean_text(r[col]) for col in ["A_EXAM", "B_TARIFF", "C_MOD", "D_QTY", "E_AMOUNT"]]

        exam = None
        for i, text in enumerate(row_texts):
            if i in [1,3,4]:
                continue
            if text and text.upper() not in GARBAGE_KEYS:
                try:
                    float(text.replace(",", "").replace("$", ""))
                    continue
                except:
                    pass
                exam = text
                break
        if not exam:
            continue

        exam_u = exam.upper()
        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        tariff_blank = pd.isna(r["B_TARIFF"]) or str(r["B_TARIFF"]).strip() in ["", "nan", "NaN", "None"]
        amt_blank = pd.isna(r["E_AMOUNT"]) or str(r["E_AMOUNT"]).strip() in ["", "nan", "NaN", "None"]
        if tariff_blank and amt_blank:
            current_subcategory = exam
            continue

        if exam_u in GARBAGE_KEYS:
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

# ---------- Template Handling ----------
def find_template_positions(ws):
    pos = {}
    header_map = {
        "DESCRIPTION": ["DESCRIPTION", "PROCEDURE", "EXAMINATION", "TEST NAME"],
        "TARIFF": ["TARIFF", "TARRIF", "RATE", "PRICE"],
        "MOD": ["MOD", "MODIFIER"],
        "QTY": ["QTY", "QUANTITY", "NO", "NUMBER"],
        "FEES": ["FEES", "CHARGE", "AMOUNT PER ITEM"]
    }
    found_headers = {key: None for key in header_map}

    for row in ws.iter_rows(min_row=1, max_row=300):
        for cell in row:
            if not cell.value:
                continue
            txt = str(cell.value).upper().strip()
            for key, variants in header_map.items():
                for v in variants:
                    if v.upper() in txt:
                        found_headers[key] = cell.column

    pos["cols"] = {k:v for k,v in found_headers.items() if v is not None}
    required = ["DESCRIPTION", "TARIFF", "MOD", "QTY", "FEES"]
    missing = [c for c in required if c not in pos["cols"]]
    if missing:
        raise ValueError(f"Missing required columns in template: {missing}")
    return pos

def write_safe(ws, r, c, value):
    if c is None:
        return
    cell = ws.cell(row=r, column=c)
    cell.value = value

def fill_excel_template(template_file, scan_rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    start_row = 22
    for idx, sr in enumerate(scan_rows):
        rowptr = start_row + idx
        write_safe(ws, rowptr, pos["cols"]["DESCRIPTION"], sr["SCAN"])
        write_safe(ws, rowptr, pos["cols"]["TARIFF"], sr["TARIFF"])
        write_safe(ws, rowptr, pos["cols"]["MOD"], sr["MODIFIER"])
        write_safe(ws, rowptr, pos["cols"]["QTY"], sr["QTY"])
        write_safe(ws, rowptr, pos["cols"]["FEES"], sr["AMOUNT"])

    # Optional: update total in template if exists
    if  "FEES" in pos["cols"]:
        total_amt = sum([r["AMOUNT"] for r in scan_rows])
        ws.cell(row=22, column=7, value=total_amt)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------- Streamlit UI ----------
st.title("Medical Quotation Generator — Template Exact Output")

uploaded_charge = st.file_uploader("Upload Charge Sheet", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template", type=["xlsx"])

if uploaded_charge:
    if st.button("Load & Parse Charge Sheet"):
        try:
            parsed = load_charge_sheet(uploaded_charge)
            st.session_state.parsed_df = parsed
            st.session_state.selected_rows = []
            st.success("Charge sheet parsed successfully.")
        except Exception as e:
            st.error(f"Failed to parse charge sheet: {e}")
            st.stop()

if "parsed_df" in st.session_state:
    df = st.session_state.parsed_df
    cats = [c for c in sorted(df["CATEGORY"].dropna().unique())] if "CATEGORY" in df.columns else []
    col1, col2 = st.columns([3,4])
    with col1:
        main_sel = st.selectbox("Select Main Category", ["-- choose --"] + cats)
    if main_sel != "-- choose --":
        subs = [s for s in sorted(df[df["CATEGORY"]==main_sel]["SUBCATEGORY"].dropna().unique())]
        sub_sel = st.selectbox("Select Subcategory", ["-- all --"] + subs) if subs else "-- all --"
        if sub_sel == "-- all --":
            sel_rows = df[df["CATEGORY"]==main_sel].to_dict(orient="records")
        else:
            sel_rows = df[(df["CATEGORY"]==main_sel)&(df["SUBCATEGORY"]==sub_sel)].to_dict(orient="records")
        st.session_state.selected_rows = sel_rows

    st.subheader("Selected Scans")
    if not st.session_state.selected_rows:
        st.info("No scans selected")
    else:
        st.dataframe(pd.DataFrame(st.session_state.selected_rows)[["SCAN","TARIFF","MODIFIER","QTY","AMOUNT"]])

        if uploaded_template and st.button("Generate Quotation (Exact Template)"):
            try:
                out = fill_excel_template(uploaded_template, st.session_state.selected_rows)
                st.download_button(
                    "Download Quotation",
                    data=out,
                    file_name="quotation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Failed to generate quotation: {e}")
