# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
import requests
from datetime import datetime, date

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# --- DEBUG / version check
st.error("RUNNING UPDATED APP.PY – VERSION 2026-01-04")

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
            st.rerun()
        else:
            st.error("Invalid credentials")
    st.stop()

# ============================================================
# HELPERS
# ============================================================
def clean_text(x):
    if pd.isna(x) or x is None:
        return ""
    return str(x).strip()

def safe_float(x, default=0.0):
    try:
        return float(x)
    except:
        return default

def safe_int(x, default=1):
    try:
        return int(float(x))
    except:
        return default

def validate_row(row):
    errors = []
    desc = clean_text(row.get("FINAL_DESC"))
    if not desc:
        errors.append("Missing description")
    try:
        if safe_int(row.get("QTY", 1)) <= 0:
            errors.append("Qty must be > 0")
    except:
        errors.append("Invalid qty")
    try:
        safe_float(row.get("TARIFF", 0))
    except:
        errors.append("Invalid tariff")
    return errors

# ============================================================
# LOAD CHARGE SHEET
# ============================================================
@st.cache_data(show_spinner=False)
def fetch_charge_sheet():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTmaRisOdFHXmFsxVA7Fx0odUq1t2QfjMvRBKqeQPgoJUdrIgSU6UhNs_-dk4jfVQ/pub?output=xlsx"
    df = pd.read_excel(url, header=None, dtype=object, engine="openpyxl")
    while df.shape[1] < 5:
        df[df.shape[1]] = None
    df = df.iloc[:, :5]
    df.columns = ["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]
    df["FINAL_DESC"] = df["SCAN"].apply(clean_text)
    df["TARIFF"] = df["TARIFF"].apply(safe_float)
    df["QTY"] = df["QTY"].apply(lambda x: safe_int(x, 1))
    df["AMOUNT"] = df["TARIFF"] * df["QTY"]
    df["IS_MAIN"] = True
    return df.reset_index(drop=True)

# ============================================================
# FETCH TEMPLATE
# ============================================================
@st.cache_data(show_spinner=False)
def fetch_template():
    url = "https://www.dropbox.com/scl/fi/iup7nwuvt5y74iu6dndak/new-template.xlsx?dl=1"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return io.BytesIO(r.content)

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
def fill_excel_template(template_file, patient, member, provider, rows, quote_date):
    wb = openpyxl.load_workbook(template_file)
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

    rowptr = pos.get("start_row", 22)
    grand_total = 0.0

    for row in rows:
        desc = clean_text(row.get("FINAL_DESC"))
        if not desc:
            desc = "INVALID DESCRIPTION"
        is_main = bool(row.get("IS_MAIN", True))
        final_desc = desc if is_main else f"   {desc}"

        write_safe(ws, rowptr, pos["cols"].get("DESCRIPTION"), final_desc)
        write_safe(ws, rowptr, pos["cols"].get("TARIFF"), row.get("TARIFF"))
        write_safe(ws, rowptr, pos["cols"].get("QTY"), row.get("QTY"))
        write_safe(ws, rowptr, pos["cols"].get("FEES"), round(row.get("AMOUNT", 0), 2))

        grand_total += row.get("AMOUNT", 0)
        rowptr += 1

    write_safe(ws, 22, pos["cols"].get("AMOUNT"), round(grand_total, 2))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# UI
# ============================================================
st.title("Medical Quotation Generator")

patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Medical Aid Provider", "CIMAS")
quote_date = st.date_input("Quotation Date", value=date.today())

df = fetch_charge_sheet()

# --- select scans
selected = st.multiselect(
    "Select scans",
    df.index,
    format_func=lambda i: df.at[i, "FINAL_DESC"]
)

rows = df.loc[selected].copy().reset_index(drop=True)

# --- add custom line
if st.button("➕ Add Custom Line Item"):
    rows = pd.concat(
        [
            rows,
            pd.DataFrame([{
                "FINAL_DESC": "ON-CALL TARIFF",
                "TARIFF": 0.0,
                "QTY": 1,
                "AMOUNT": 0.0,
                "IS_MAIN": True
            }])
        ],
        ignore_index=True
    )

# --- editor
if not rows.empty:
    st.subheader("Edit Final Descriptions & Tariffs")
    editor = st.data_editor(
        rows,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "FINAL_DESC": st.column_config.TextColumn("Final Description"),
            "TARIFF": st.column_config.NumberColumn("Tariff"),
            "QTY": st.column_config.NumberColumn("Qty", min_value=1),
        }
    )

    editor["AMOUNT"] = editor["TARIFF"] * editor["QTY"]

    # warn if blank
    valid_rows = editor[editor["FINAL_DESC"].astype(str).str.strip() != ""]
    if len(valid_rows) != len(editor):
        st.error("Some rows have blank descriptions and will be excluded from the quotation.")

    st.metric("Grand Total", f"${valid_rows['AMOUNT'].sum():,.2f}")

    # --- download
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
            "Download Quotation",
            out,
            "quotation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
