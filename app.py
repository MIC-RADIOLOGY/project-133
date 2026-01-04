# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
import requests
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
            st.success("Login successful.")
        else:
            st.error("Invalid credentials")
    st.stop()

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
COMPONENT_KEYS = {"PELVIS", "CONSUMABLES", "FF", "IV", "IV CONTRAST", "IV CONTRAST 100MLS"}
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
    except:
        return default

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "").strip())
    except:
        return default

# ------------------------------------------------------------
# PARSER
# ------------------------------------------------------------
def load_charge_sheet(file):
    df_raw = pd.read_excel(file, header=None, dtype=object, engine="openpyxl")
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
            MAIN_CATEGORIES.add(exam_u)
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
            "TARIFF": safe_float(r["B_TARIFF"], 0.0),
            "MODIFIER": clean_text(r["C_MOD"]),
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
    try:
        ws.cell(row=r, column=c).value = value
    except:
        for mr in ws.merged_cells.ranges:
            if ws.cell(row=r, column=c).coordinate in mr:
                ws.cell(row=mr.min_row, column=mr.min_col).value = value
                return

def append_after_label(ws, r, c, value):
    if not value:
        return
    cell = ws.cell(row=r, column=c)
    cell.value = f"{cell.value or ''} {value}".strip()

def write_below_label(ws, r, c, value):
    target = ws.cell(row=r + 1, column=c)
    try:
        target.value = value
    except:
        for mr in ws.merged_cells.ranges:
            if target.coordinate in mr:
                ws.cell(row=mr.min_row, column=mr.min_col).value = value

def find_template_positions(ws):
    pos = {}
    headers = ["DESCRIPTION", "TARIFF", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]

    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if not cell.value:
                continue
            t = str(cell.value).upper()

            if "PATIENT" in t:
                pos["patient_cell"] = (cell.row, cell.column)
            if "MEMBER" in t:
                pos["member_cell"] = (cell.row, cell.column)
            if "PROVIDER" in t or "MEDICAL AID" in t:
                pos["provider_cell"] = (cell.row, cell.column)
            if t.strip() == "DATE":
                pos["date_cell"] = (cell.row, cell.column)

            if any(h in t for h in headers):
                pos.setdefault("cols", {})
                pos["table_start_row"] = cell.row + 1
                for h in headers:
                    if h in t:
                        pos["cols"][h] = cell.column
    return pos

def fill_excel_template(template_file, patient, member, provider, scan_rows, date_value=None):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    # ---- HARD FALLBACKS (CRITICAL FIX) ----
    cols = pos.get("cols", {})
    cols.setdefault("DESCRIPTION", 1)
    cols.setdefault("TARIFF", 2)
    cols.setdefault("MOD", 3)
    cols.setdefault("QTY", 4)
    cols.setdefault("FEES", 5)
    cols.setdefault("AMOUNT", 5)

    if "patient_cell" in pos:
        append_after_label(ws, *pos["patient_cell"], patient)
    if "member_cell" in pos:
        append_after_label(ws, *pos["member_cell"], member)
    if "provider_cell" in pos:
        append_after_label(ws, *pos["provider_cell"], provider)
    if "date_cell" in pos and date_value:
        write_below_label(ws, *pos["date_cell"], date_value.strftime("%d/%m/%Y"))

    rowptr = pos.get("table_start_row", 22)
    grand_total = 0.0

    for sr in scan_rows:
        scan_desc = str(sr.get("SCAN", "")).strip()
        tariff = sr.get("TARIFF", 0.0)
        qty = sr.get("QTY", 1)
        modifier = sr.get("MODIFIER", "")
        amount = round((tariff or 0) * (qty or 1), 2)

        write_safe(ws, rowptr, cols["DESCRIPTION"], scan_desc)
        write_safe(ws, rowptr, cols["TARIFF"], tariff)
        write_safe(ws, rowptr, cols["MOD"], modifier)
        write_safe(ws, rowptr, cols["QTY"], qty)
        write_safe(ws, rowptr, cols["FEES"], amount)

        grand_total += amount
        rowptr += 1

    write_safe(ws, rowptr + 1, cols["AMOUNT"], round(grand_total, 2))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ------------------------------------------------------------
# LOAD FILES
# ------------------------------------------------------------
@st.cache_data(show_spinner=False)
def fetch_charge_sheet():
    url = (
        "https://docs.google.com/spreadsheets/d/e/"
        "2PACX-1vTmaRisOdFHXmFsxVA7Fx0odUq1t2QfjMvRBKqeQPgoJUdrIgSU6UhNs_-dk4jfVQ"
        "/pub?output=xlsx"
    )
    return load_charge_sheet(url)

@st.cache_data(show_spinner=False)
def fetch_quote_template():
    url = "https://www.dropbox.com/scl/fi/iup7nwuvt5y74iu6dndak/new-template.xlsx?dl=1"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return io.BytesIO(r.content)

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")
quotation_date = st.date_input("Quotation Date", datetime.today())

if "df" not in st.session_state:
    st.session_state.df = fetch_charge_sheet()

df = st.session_state.df

main_sel = st.selectbox("Select Main Category", sorted(df["CATEGORY"].unique()))
scans = df[df["CATEGORY"] == main_sel].reset_index(drop=True)

selected = st.multiselect(
    "Select scans",
    scans.index.tolist(),
    format_func=lambda i: scans.at[i, "SCAN"]
)

selected_rows = [scans.iloc[i].to_dict() for i in selected]

if st.button("➕ Add Custom Line Item"):
    selected_rows.append({
        "CATEGORY": "CUSTOM",
        "SUBCATEGORY": None,
        "SCAN": "ON-CALL TARIFF",
        "TARIFF": 0.0,
        "QTY": 1,
        "MODIFIER": "",
        "AMOUNT": 0.0
    })

if selected_rows:
    edit_df = pd.DataFrame(selected_rows)

    edit_df = st.data_editor(
        edit_df,
        column_config={
            "SCAN": st.column_config.TextColumn("Final Description"),
            "TARIFF": st.column_config.NumberColumn("Tariff"),
            "QTY": st.column_config.NumberColumn("Qty", min_value=1),
        },
        use_container_width=True
    )

    if st.button("Generate & Download Quotation"):
        out = fill_excel_template(
            fetch_quote_template(),
            patient,
            member,
            provider,
            edit_df.to_dict("records"),
            quotation_date
        )
        st.download_button(
            "Download Quotation",
            data=out,
            file_name="quotation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
