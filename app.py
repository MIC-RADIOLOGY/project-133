import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

# ================= PAGE CONFIG =================
st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ================= CONFIG =================
MAIN_CATEGORIES = {
    "ULTRA SOUND", "ULTRASOUND", "CT SCAN", "X-RAY", "XRAY", "FLUROSCOPY"
}

GARBAGE_KEYS = {
    "TOTAL", "SUB TOTAL", "GRAND TOTAL",
    "CO", "CO PAY", "CO-PAY", "CO PAYMENT"
}

# ================= UTILITIES =================
def clean(x):
    if pd.isna(x) or x is None:
        return ""
    return str(x).replace("\xa0", " ").strip()

def norm(x):
    return clean(x).upper()

def safe_int(x, default=1):
    try:
        return int(float(str(x).replace(",", "").strip()))
    except (ValueError, TypeError):
        return default

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "").strip())
    except (ValueError, TypeError):
        return default

# ================= PARSER =================
@st.cache_data(show_spinner=False)
def load_charge_sheet(file):
    df = pd.read_excel(file, header=None, dtype=object)

    # Ensure minimum columns
    while df.shape[1] < 5:
        df[df.shape[1]] = None

    df = df.iloc[:, :5]
    df.columns = ["EXAM", "TARIFF", "MOD", "QTY", "AMOUNT"]

    rows = []
    current_category = None
    current_subcategory = None

    for _, r in df.iterrows():
        exam = clean(r.EXAM)
        if not exam:
            continue

        exam_u = exam.upper()

        if any(k in exam_u for k in GARBAGE_KEYS):
            continue

        # Detect main category
        if any(cat in exam_u for cat in MAIN_CATEGORIES):
            current_category = exam
            current_subcategory = None
            continue

        # Detect subcategory (no numeric data)
        if (
            clean(r.AMOUNT) == "" and
            clean(r.QTY) == "" and
            not any(c.isdigit() for c in exam)
        ):
            current_subcategory = exam
            continue

        qty = safe_int(r.QTY, 1)
        tariff = safe_float(r.TARIFF, 0.0)
        amount = tariff * qty

        rows.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "TARIFF": tariff,
            "MODIFIER": clean(r.MOD),
            "QTY": qty,
            "AMOUNT": amount
        })

    return pd.DataFrame(rows)

# ================= EXCEL WRITER =================
def fill_excel_template(template_file, patient, member, provider, rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    col_map = {}
    start_row = None

    # Locate table headers
    for row in ws.iter_rows(max_row=200):
        for cell in row:
            if not cell.value:
                continue
            t = norm(cell.value)

            if "DESCRIPTION" in t:
                col_map["SCAN"] = cell.column
                start_row = cell.row + 1
            elif "TARIFF" in t or "TARRIF" in t:
                col_map["TARIFF"] = cell.column
            elif "MOD" in t:
                col_map["MODIFIER"] = cell.column
            elif "QTY" in t or "QUANTITY" in t:
                col_map["QTY"] = cell.column
            elif "AMOUNT" in t or "FEES" in t:
                col_map["AMOUNT"] = cell.column

    # Write patient info (best-effort)
    for row in ws.iter_rows(max_row=50):
        for cell in row:
            if not cell.value:
                continue
            t = norm(cell.value)
            if "PATIENT" in t:
                ws.cell(cell.row, cell.column + 1).value = patient
            elif "MEMBER" in t:
                ws.cell(cell.row, cell.column + 1).value = member
            elif "PROVIDER" in t or "MEDICAL AID" in t:
                ws.cell(cell.row, cell.column + 1).value = provider
            elif t == "DATE":
                ws.cell(cell.row + 1, cell.column).value = datetime.today().strftime("%d/%m/%Y")

    # Write scan rows
    r = start_row
    for item in rows:
        ws.cell(r, col_map["SCAN"]).value = item["SCAN"]
        ws.cell(r, col_map["TARIFF"]).value = item["TARIFF"]
        ws.cell(r, col_map["MODIFIER"]).value = item["MODIFIER"]
        ws.cell(r, col_map["QTY"]).value = item["QTY"]
        ws.cell(r, col_map["AMOUNT"]).value = item["AMOUNT"]
        r += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ================= STREAMLIT UI =================
st.title("Medical Quotation Generator")

charge_file = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
template_file = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if charge_file and st.button("Load & Parse Charge Sheet"):
    st.session_state.df = load_charge_sheet(charge_file)
    st.success("Charge sheet parsed successfully")

if "df" in st.session_state:
    df = st.session_state.df

    categories = sorted(df["CATEGORY"].dropna().unique())
    selected_categories = st.multiselect("Select Categories", categories)

    filtered = df[df["CATEGORY"].isin(selected_categories)] if selected_categories else df

    st.subheader("Select & Edit Scans")
    edited_df = st.data_editor(
        filtered,
        use_container_width=True,
        num_rows="fixed"
    )

    edited_df["AMOUNT"] = edited_df["TARIFF"] * edited_df["QTY"]

    total = edited_df["AMOUNT"].sum()
    st.metric("Total Quotation Amount", f"{total:,.2f}")

    st.subheader("Quotation Preview")
    st.dataframe(edited_df)

    if template_file and st.button("Generate & Download Quotation"):
        output = fill_excel_template(
            template_file,
            patient,
            member,
            provider,
            edited_df.to_dict("records")
        )

        st.download_button(
            "Download Quotation",
            data=output,
            file_name=f"quotation_{patient or 'patient'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
