# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
import requests
from datetime import datetime
from openpyxl.styles import Font

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
            st.rerun()
        else:
            st.error("Invalid credentials")
    st.stop()

# ============================================================
# HELPERS
# ============================================================
ERROR_FONT = Font(color="FF0000", bold=True)

def safe_int(x, default=1):
    try:
        return int(float(x))
    except:
        return default

def safe_float(x, default=0.0):
    try:
        return float(x)
    except:
        return default

def normalize_text(x):
    if x is None or pd.isna(x):
        return ""
    return str(x).strip()

def validate_row(row):
    errors = []

    desc = normalize_text(row.get("SCAN"))
    qty = row.get("QTY")
    tariff = row.get("TARIFF")

    if not desc:
        errors.append("Missing description")

    try:
        qty = int(qty)
        if qty <= 0:
            errors.append("Qty must be > 0")
    except:
        errors.append("Invalid qty")

    try:
        float(tariff)
    except:
        errors.append("Invalid tariff")

    return errors

# ============================================================
# LOAD CHARGE SHEET
# ============================================================
@st.cache_data(show_spinner=False)
def fetch_charge_sheet():
    url = (
        "https://docs.google.com/spreadsheets/d/e/"
        "2PACX-1vTmaRisOdFHXmFsxVA7Fx0odUq1t2QfjMvRBKqeQPgoJUdrIgSU6UhNs_-dk4jfVQ"
        "/pub?output=xlsx"
    )

    df = pd.read_excel(url, header=None)
    df = df.iloc[:, :3]
    df.columns = ["SCAN", "TARIFF", "QTY"]

    df = df.dropna(subset=["SCAN"])
    df["TARIFF"] = df["TARIFF"].fillna(0)
    df["QTY"] = df["QTY"].fillna(1)
    df["AMOUNT"] = df["TARIFF"] * df["QTY"]
    df["IS_MAIN_SCAN"] = True

    return df.reset_index(drop=True)

@st.cache_data(show_spinner=False)
def fetch_template():
    url = "https://www.dropbox.com/scl/fi/iup7nwuvt5y74iu6dndak/new-template.xlsx?dl=1"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return io.BytesIO(r.content)

# ============================================================
# EXCEL EXPORT (CRASH-PROOF)
# ============================================================
def fill_excel_template(template_file, rows, patient, member, provider, qdate):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    ws["A5"].value = f"Patient: {patient}"
    ws["A6"].value = f"Member: {member}"
    ws["A7"].value = f"Provider: {provider}"
    ws["A8"].value = f"Date: {qdate.strftime('%d/%m/%Y')}"

    rowptr = 22
    grand_total = 0.0

    for row in rows:
        errors = validate_row(row)

        scan_desc = normalize_text(row.get("SCAN"))
        is_main = bool(row.get("IS_MAIN_SCAN", True))
        qty = safe_int(row.get("QTY"), 1)
        tariff = safe_float(row.get("TARIFF"), 0.0)
        amount = round(qty * tariff, 2)

        # ---- INVALID ROW → RED WARNING, NO CRASH ----
        if errors:
            cell = ws.cell(row=rowptr, column=1)
            cell.value = f"⚠ INVALID ROW: {', '.join(errors)}"
            cell.font = ERROR_FONT
            rowptr += 1
            continue

        if not scan_desc:
            scan_desc = "UNNAMED SCAN"

        # ---- SAFE INDENT (NO STRING CONCAT WITH NON-STRING) ----
        final_desc = scan_desc if is_main else f"   {scan_desc}"

        ws.cell(row=rowptr, column=1).value = final_desc
        ws.cell(row=rowptr, column=2).value = tariff
        ws.cell(row=rowptr, column=3).value = qty
        ws.cell(row=rowptr, column=4).value = amount

        grand_total += amount
        rowptr += 1

    ws.cell(row=rowptr + 1, column=4).value = round(grand_total, 2)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# UI
# ============================================================
st.title("Medical Quotation Generator")

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")
quotation_date = st.date_input("Quotation Date", datetime.today())

if "charge_df" not in st.session_state:
    st.session_state.charge_df = fetch_charge_sheet()

if "edit_df" not in st.session_state:
    st.session_state.edit_df = pd.DataFrame()

df = st.session_state.charge_df

selected = st.multiselect(
    "Select scans",
    df.index.tolist(),
    format_func=lambda i: df.at[i, "SCAN"]
)

# Load selection ONCE
if selected and st.session_state.edit_df.empty:
    st.session_state.edit_df = df.loc[selected].copy().reset_index(drop=True)

# Add custom line
if st.button("➕ Add Custom Line Item"):
    st.session_state.edit_df = pd.concat(
        [
            st.session_state.edit_df,
            pd.DataFrame([{
                "SCAN": "ON-CALL TARIFF",
                "TARIFF": 0.0,
                "QTY": 1,
                "AMOUNT": 0.0,
                "IS_MAIN_SCAN": True
            }])
        ],
        ignore_index=True
    )

# Editor
if not st.session_state.edit_df.empty:
    st.subheader("Edit Final Quotation (Printed to Excel)")

    st.session_state.edit_df = st.data_editor(
        st.session_state.edit_df,
        column_config={
            "SCAN": st.column_config.TextColumn("Final Description"),
            "TARIFF": st.column_config.NumberColumn("Tariff"),
            "QTY": st.column_config.NumberColumn("Qty", min_value=1),
            "IS_MAIN_SCAN": st.column_config.CheckboxColumn("Main Item"),
        },
        use_container_width=True
    )

    st.session_state.edit_df["AMOUNT"] = (
        st.session_state.edit_df["TARIFF"].apply(safe_float)
        * st.session_state.edit_df["QTY"].apply(safe_int)
    )

    st.metric(
        "Grand Total",
        f"${st.session_state.edit_df['AMOUNT'].sum():,.2f}"
    )

    if st.button("Generate & Download Quotation"):
        out = fill_excel_template(
            fetch_template(),
            st.session_state.edit_df.to_dict("records"),
            patient,
            member,
            provider,
            quotation_date
        )
        st.download_button(
            "Download Quotation",
            data=out,
            file_name="quotation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
