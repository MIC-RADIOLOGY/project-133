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
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u == "admin" and p == "Jamela2003":
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("Invalid credentials")
    st.stop()

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def safe_float(x, d=0.0):
    try:
        return float(x)
    except:
        return d

def safe_int(x, d=1):
    try:
        return int(x)
    except:
        return d

# ------------------------------------------------------------
# LOAD CHARGE SHEET
# ------------------------------------------------------------
@st.cache_data(show_spinner=False)
def fetch_charge_sheet():
    url = (
        "https://docs.google.com/spreadsheets/d/e/"
        "2PACX-1vTmaRisOdFHXmFsxVA7Fx0odUq1t2QfjMvRBKqeQPgoJUdrIgSU6UhNs_-dk4jfVQ"
        "/pub?output=xlsx"
    )
    df = pd.read_excel(url, header=None)
    df.columns = ["SCAN", "TARIFF", "MOD", "QTY", "AMOUNT"]
    df = df.dropna(subset=["SCAN"])
    df["QTY"] = df["QTY"].fillna(1)
    df["TARIFF"] = df["TARIFF"].fillna(0)
    df["AMOUNT"] = df["TARIFF"] * df["QTY"]
    return df

# ------------------------------------------------------------
# EXCEL EXPORT
# ------------------------------------------------------------
def export_excel(template_file, rows, patient, member, provider, qdate):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    row = 22
    total = 0

    for r in rows:
        ws.cell(row=row, column=1).value = r["SCAN"]
        ws.cell(row=row, column=2).value = r["TARIFF"]
        ws.cell(row=row, column=3).value = r["QTY"]
        ws.cell(row=row, column=4).value = r["AMOUNT"]
        total += r["AMOUNT"]
        row += 1

    ws.cell(row=row + 1, column=4).value = round(total, 2)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

@st.cache_data(show_spinner=False)
def fetch_template():
    url = "https://www.dropbox.com/scl/fi/iup7nwuvt5y74iu6dndak/new-template.xlsx?dl=1"
    r = requests.get(url)
    return io.BytesIO(r.content)

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member")
provider = st.text_input("Provider", value="CIMAS")
qdate = st.date_input("Quotation Date", datetime.today())

if "charge_df" not in st.session_state:
    st.session_state.charge_df = fetch_charge_sheet()

if "edit_df" not in st.session_state:
    st.session_state.edit_df = pd.DataFrame()

df = st.session_state.charge_df

# SELECT SCANS
selected = st.multiselect(
    "Select scans",
    df.index.tolist(),
    format_func=lambda i: df.at[i, "SCAN"]
)

# LOAD SELECTION INTO EDIT STATE (ONLY ONCE)
if selected:
    base = df.loc[selected].copy()
    base["AMOUNT"] = base["TARIFF"] * base["QTY"]
    st.session_state.edit_df = base.reset_index(drop=True)

# ADD CUSTOM ROW
if st.button("➕ Add Custom Line Item"):
    st.session_state.edit_df = pd.concat(
        [
            st.session_state.edit_df,
            pd.DataFrame([{
                "SCAN": "ON-CALL TARIFF",
                "TARIFF": 0.0,
                "QTY": 1,
                "AMOUNT": 0.0
            }])
        ],
        ignore_index=True
    )

# EDITOR (THIS IS NOW PERSISTENT)
if not st.session_state.edit_df.empty:
    st.subheader("Edit Final Quotation (THIS WILL PRINT)")
    st.session_state.edit_df = st.data_editor(
        st.session_state.edit_df,
        column_config={
            "SCAN": st.column_config.TextColumn("Final Description"),
            "TARIFF": st.column_config.NumberColumn("Tariff"),
            "QTY": st.column_config.NumberColumn("Qty", min_value=1),
        },
        use_container_width=True
    )

    # RECALCULATE
    st.session_state.edit_df["AMOUNT"] = (
        st.session_state.edit_df["TARIFF"].apply(safe_float)
        * st.session_state.edit_df["QTY"].apply(safe_int)
    )

    st.metric(
        "Grand Total",
        f"${st.session_state.edit_df['AMOUNT'].sum():,.2f}"
    )

    if st.button("Generate & Download Quotation"):
        out = export_excel(
            fetch_template(),
            st.session_state.edit_df.to_dict("records"),
            patient,
            member,
            provider,
            qdate
        )
        st.download_button(
            "Download Quotation",
            data=out,
            file_name="quotation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
