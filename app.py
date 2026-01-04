# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
import requests
from datetime import datetime, date

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ============================================================
# LOGIN
# ============================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("Login Required")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username == "admin" and password == "Jamela2003":
            st.session_state.logged_in = True
            st.success("Login successful")
        else:
            st.error("Invalid credentials")
    st.stop()

# ============================================================
# CONSTANTS
# ============================================================
COMPONENT_KEYS = {"PELVIS", "CONSUMABLES", "FF", "IV", "IV CONTRAST", "IV CONTRAST 100MLS"}
GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# ============================================================
# HELPERS
# ============================================================
def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "").strip())
    except:
        return default

def safe_int(x, default=1):
    try:
        return int(float(str(x).replace(",", "").strip()))
    except:
        return default

# ============================================================
# LOAD CHARGE SHEET
# ============================================================
def load_charge_sheet(source):
    df = pd.read_excel(source, header=None, dtype=object, engine="openpyxl")

    while df.shape[1] < 5:
        df[df.shape[1]] = None

    df = df.iloc[:, :5]
    df.columns = ["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]

    rows = []
    current_category = None
    current_subcategory = None

    for _, r in df.iterrows():
        scan = clean_text(r["SCAN"])
        if not scan:
            continue

        scan_u = scan.upper()

        if scan_u.endswith("SCAN") or scan_u in {"XRAY", "MRI", "CT", "ULTRASOUND"}:
            current_category = scan
            current_subcategory = None
            continue

        if scan_u in GARBAGE_KEYS:
            continue

        if clean_text(r["TARIFF"]) == "" and clean_text(r["AMOUNT"]) == "":
            current_subcategory = scan
            continue

        if not current_category:
            continue

        rows.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "FINAL_DESC": scan,
            "IS_MAIN": scan_u not in COMPONENT_KEYS,
            "TARIFF": safe_float(r["TARIFF"]),
            "QTY": safe_int(r["QTY"]),
            "AMOUNT": safe_float(r["AMOUNT"])
        })

    return pd.DataFrame(rows)

# ============================================================
# EXCEL HELPERS
# ============================================================
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
    cell = ws.cell(row=r, column=c)
    existing = str(cell.value) if cell.value else ""
    cell.value = f"{existing} {value}".strip()

def write_below_label(ws, r, c, value):
    ws.cell(row=r + 1, column=c).value = value

def find_positions(ws):
    pos = {"cols": {}}
    headers = ["DESCRIPTION", "TARIFF", "MOD", "QTY", "FEES", "AMOUNT"]

    for row in ws.iter_rows(max_row=200):
        for cell in row:
            if not cell.value:
                continue
            t = str(cell.value).upper()

            if "PATIENT" in t:
                pos["patient"] = (cell.row, cell.column)
            if "MEMBER" in t:
                pos["member"] = (cell.row, cell.column)
            if "PROVIDER" in t or "MEDICAL AID" in t:
                pos["provider"] = (cell.row, cell.column)
            if t.strip() == "DATE":
                pos["date"] = (cell.row, cell.column)

            for h in headers:
                if h in t:
                    pos["cols"][h] = cell.column
                    pos["start_row"] = cell.row + 1

    return pos

# ============================================================
# FILL TEMPLATE (CRASH-PROOF)
# ============================================================
def fill_excel_template(template, patient, member, provider, rows, quote_date):
    wb = openpyxl.load_workbook(template)
    ws = wb.active
    pos = find_positions(ws)

    if "patient" in pos:
        append_after_label(ws, *pos["patient"], patient)
    if "member" in pos:
        append_after_label(ws, *pos["member"], member)
    if "provider" in pos:
        append_after_label(ws, *pos["provider"], provider)
    if "date" in pos:
        write_below_label(ws, *pos["date"], quote_date.strftime("%d/%m/%Y"))

    r = pos.get("start_row", 22)
    total = 0.0

    for row in rows:
        desc = str(row.get("FINAL_DESC") or "").strip()
        if not desc:
            desc = "INVALID DESCRIPTION"

        is_main = bool(row.get("IS_MAIN", True))
        final_desc = desc if is_main else f"   {desc}"

        write_safe(ws, r, pos["cols"].get("DESCRIPTION"), final_desc)
        write_safe(ws, r, pos["cols"].get("TARIFF"), row.get("TARIFF"))
        write_safe(ws, r, pos["cols"].get("QTY"), row.get("QTY"))
        write_safe(ws, r, pos["cols"].get("FEES"), round(row.get("AMOUNT", 0), 2))

        total += row.get("AMOUNT", 0)
        r += 1

    write_safe(ws, 22, pos["cols"].get("AMOUNT"), round(total, 2))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# LOAD FILES
# ============================================================
@st.cache_data
def fetch_charge_sheet():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTmaRisOdFHXmFsxVA7Fx0odUq1t2QfjMvRBKqeQPgoJUdrIgSU6UhNs_-dk4jfVQ/pub?output=xlsx"
    return load_charge_sheet(url)

@st.cache_data
def fetch_template():
    url = "https://www.dropbox.com/scl/fi/iup7nwuvt5y74iu6dndak/new-template.xlsx?dl=1"
    r = requests.get(url)
    return io.BytesIO(r.content)

# ============================================================
# UI
# ============================================================
st.title("Medical Quotation Generator")

patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Medical Aid Provider", "CIMAS")
quote_date = st.date_input("Quotation Date", value=date.today())

df = fetch_charge_sheet()

category = st.selectbox("Category", sorted(df["CATEGORY"].unique()))
subset = df[df["CATEGORY"] == category].reset_index(drop=True)

selected = st.multiselect(
    "Select scans",
    subset.index,
    format_func=lambda i: subset.at[i, "FINAL_DESC"]
)

rows = subset.loc[selected].copy()

st.subheader("Edit Final Descriptions & Add Custom Rows")

editor = st.data_editor(
    rows,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "FINAL_DESC": st.column_config.TextColumn("Final Description"),
        "TARIFF": st.column_config.NumberColumn("Tariff"),
        "AMOUNT": st.column_config.NumberColumn("Amount")
    }
)

valid_rows = editor[editor["FINAL_DESC"].astype(str).str.strip() != ""]

if len(valid_rows) != len(editor):
    st.error("Some rows have blank descriptions and will be excluded.")

st.metric("Grand Total", f"${valid_rows['AMOUNT'].sum():,.2f}")

if st.button("Generate Quotation"):
    out = fill_excel_template(
        fetch_template(),
        patient,
        member,
        provider,
        valid_rows.to_dict("records"),
        quote_date
    )

    st.download_button(
        "Download Excel",
        out,
        "quotation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
