import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

# ================= CONFIG =================
st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

MAIN_CATEGORIES = {
    "ULTRA SOUND", "ULTRASOUND", "CT SCAN", "X-RAY", "XRAY", "FLUROSCOPY"
}

GARBAGE_KEYS = {
    "TOTAL", "SUB TOTAL", "GRAND TOTAL",
    "CO", "CO PAY", "CO-PAY", "CO PAYMENT"
}

# ================= UTILITIES =================
def clean(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def norm(x):
    return clean(x).upper()

def safe_int(x, default=1):
    try:
        return int(float(str(x).replace(",", "")))
    except:
        return default

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "")))
    except:
        return default

# ================= PARSER =================
def load_charge_sheet(file):
    df = pd.read_excel(file, header=None, dtype=object)
    while df.shape[1] < 5:
        df[df.shape[1]] = None

    df = df.iloc[:, :5]
    df.columns = ["EXAM", "TARIFF", "MOD", "QTY", "AMOUNT"]

    rows = []
    current_cat = None
    current_sub = None

    for _, r in df.iterrows():
        exam = clean(r.EXAM)
        if not exam:
            continue

        exam_u = exam.upper()

        if any(k in exam_u for k in GARBAGE_KEYS):
            continue

        if any(cat in exam_u for cat in MAIN_CATEGORIES):
            current_cat = exam
            current_sub = None
            continue

        # Subcategory heuristic
        if (
            clean(r.AMOUNT) == "" and
            clean(r.QTY) == "" and
            not any(c.isdigit() for c in exam)
        ):
            current_sub = exam
            continue

        qty = safe_int(r.QTY, 1)
        tariff = safe_float(r.TARIFF, 0.0)
        amount = tariff * qty

        rows.append({
            "CATEGORY": current_cat,
            "SUBCATEGORY": current_sub,
            "SCAN": exam,
            "TARIFF": tariff,
            "MODIFIER": clean(r.MOD),
            "QTY": qty,
            "AMOUNT": amount
        })

    return pd.DataFrame(rows)

# ================= EXCEL WRITER =================
def fill_excel_template(template, patient, member, provider, rows):
    wb = openpyxl.load_workbook(template)
    ws = wb.active

    # Find headers
    col_map = {}
    start_row = None

    for row in ws.iter_rows(max_row=200):
        for cell in row:
            if not cell.value:
                continue
            t = norm(cell.value)

            if "DESCRIPTION" in t:
                col_map["SCAN"] = cell.column
                start_row = cell.row + 1
            elif "TARRIF" in t or "TARIFF" in t:
                col_map["TARIFF"] = cell.column
            elif "MOD" in t:
                col_map["MODIFIER"] = cell.column
            elif "QTY" in t:
                col_map["QTY"] = cell.column
            elif "AMOUNT" in t or "FEES" in t:
                col_map["AMOUNT"] = cell.column

    r = start_row
    for row in rows:
        ws.cell(r, col_map["SCAN"]).value = row["SCAN"]
        ws.cell(r, col_map["TARIFF"]).value = row["TARIFF"]
        ws.cell(r, col_map["MODIFIER"]).value = row["MODIFIER"]
        ws.cell(r, col_map["QTY"]).value = row["QTY"]
        ws.cell(r, col_map["AMOUNT"]).value = row["AMOUNT"]
        r += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ================= STREAMLIT UI =================
st.title("Medical Quotation Generator")

charge_file = st.file_uploader("Upload Charge Sheet", ["xlsx"])
template_file = st.file_uploader("Upload Excel Template", ["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Medical Aid Provider", "CIMAS")

if charge_file and st.button("Parse Charge Sheet"):
    st.session_state.df = load_charge_sheet(charge_file)
    st.success("Parsed successfully")

if "df" in st.session_state:
    df = st.session_state.df

    cats = sorted(df.CATEGORY.dropna().unique())
    selected_cats = st.multiselect("Select Categories", cats)

    filtered = df[df.CATEGORY.isin(selected_cats)] if selected_cats else df

    st.subheader("Select & Edit Scans")
    edited = st.data_editor(
        filtered,
        use_container_width=True,
        num_rows="fixed"
    )

    edited["AMOUNT"] = edited["TARIFF"] * edited["QTY"]

    total = edited["AMOUNT"].sum()
    st.metric("Total Amount", f"{total:,.2f}")

    st.subheader("Quotation Preview")
    st.dataframe(edited)

    if template_file and st.button("Generate Quotation"):
        out = fill_excel_template(
            template_file, patient, member, provider,
            edited.to_dict("records")
        )
        st.download_button(
            "Download Quotation",
            data=out,
            file_name=f"quotation_{patient or 'patient'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
