import streamlit as st
import pandas as pd
import openpyxl
import io
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
            st.success("Login successful")
        else:
            st.error("Invalid credentials")
    st.stop()

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
COMPONENT_KEYS = {"PELVIS", "CONSUMABLES", "FF", "IV", "IV CONTRAST"}
GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO", ""}
MAIN_CATEGORIES = set()

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def safe_int(x, d=1):
    try:
        return int(float(str(x).replace(",", "")))
    except:
        return d

def safe_float(x, d=0.0):
    try:
        return float(str(x).replace(",", "")))
    except:
        return d

# ------------------------------------------------------------
# PARSER
# ------------------------------------------------------------
def load_charge_sheet(file):
    df = pd.read_excel(file, header=None, dtype=object)
    while df.shape[1] < 5:
        df[df.shape[1]] = None

    df = df.iloc[:, :5]
    df.columns = ["EXAM", "TARIFF", "MOD", "QTY", "AMOUNT"]

    rows = []
    cat = None
    sub = None

    for _, r in df.iterrows():
        exam = clean_text(r["EXAM"])
        if not exam:
            continue

        eu = exam.upper()

        if eu.endswith("SCAN") or eu in {"XRAY", "MRI", "ULTRASOUND"}:
            cat = exam
            sub = None
            continue

        if eu in GARBAGE_KEYS:
            continue

        if clean_text(r["TARIFF"]) == "" and clean_text(r["AMOUNT"]) == "":
            sub = exam
            continue

        if not cat:
            continue

        rows.append({
            "CATEGORY": cat,
            "SUBCATEGORY": sub,
            "SCAN": exam,
            "IS_MAIN_SCAN": eu not in COMPONENT_KEYS,
            "TARIFF": safe_float(r["TARIFF"]),
            "MODIFIER": clean_text(r["MOD"]),
            "QTY": safe_int(r["QTY"]),
            "AMOUNT": safe_float(r["AMOUNT"])
        })

    return pd.DataFrame(rows)

# ------------------------------------------------------------
# EXCEL POSITION DETECTION (FIXED)
# ------------------------------------------------------------
def find_positions(ws):
    pos = {
        "COLS": {},
        "MOD_COL": None
    }

    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if not cell.value:
                continue

            txt = str(cell.value).upper()

            if "DESCRIPTION" in txt:
                pos["COLS"]["DESC"] = cell.column
                pos["START_ROW"] = cell.row + 1

            if "TARIFF" in txt:
                pos["COLS"]["TARIFF"] = cell.column

            if "QTY" in txt:
                pos["COLS"]["QTY"] = cell.column

            if "FEE" in txt:
                pos["COLS"]["FEES"] = cell.column

            if "AMOUNT" in txt:
                pos["COLS"]["AMOUNT"] = cell.column

            # 🔥 GUARANTEED MOD DETECTION
            if "MOD" in txt:
                pos["MOD_COL"] = cell.column

            if "PATIENT" in txt:
                pos["PATIENT"] = (cell.row, cell.column)

            if "MEMBER" in txt:
                pos["MEMBER"] = (cell.row, cell.column)

            if "PROVIDER" in txt or "MEDICAL AID" in txt:
                pos["PROVIDER"] = (cell.row, cell.column)

            if txt.strip() == "DATE":
                pos["DATE"] = (cell.row, cell.column)

    return pos

# ------------------------------------------------------------
# SAFE WRITE
# ------------------------------------------------------------
def write(ws, r, c, v):
    if not c:
        return
    try:
        ws.cell(r, c).value = v
    except:
        for m in ws.merged_cells.ranges:
            if ws.cell(r, c).coordinate in m:
                ws.cell(m.min_row, m.min_col).value = v
                return

# ------------------------------------------------------------
# TEMPLATE FILL
# ------------------------------------------------------------
def fill_template(tpl, patient, member, provider, rows):
    wb = openpyxl.load_workbook(tpl)
    ws = wb.active
    p = find_positions(ws)

    if "PATIENT" in p:
        write(ws, *p["PATIENT"], patient)
    if "MEMBER" in p:
        write(ws, *p["MEMBER"], member)
    if "PROVIDER" in p:
        write(ws, *p["PROVIDER"], provider)
    if "DATE" in p:
        write(ws, p["DATE"][0] + 1, p["DATE"][1],
              datetime.today().strftime("%d/%m/%Y"))

    r = p.get("START_ROW", 22)
    total = 0

    for s in rows:
        write(ws, r, p["COLS"]["DESC"],
              s["SCAN"] if s["IS_MAIN_SCAN"] else "   " + s["SCAN"])
        write(ws, r, p["COLS"].get("TARIFF"), s["TARIFF"])
        write(ws, r, p["MOD_COL"], s["MODIFIER"])   # ✅ FIX
        write(ws, r, p["COLS"].get("QTY"), s["QTY"])
        write(ws, r, p["COLS"].get("FEES"),
              round(s["AMOUNT"] / s["QTY"], 2))
        total += s["AMOUNT"]
        r += 1

    write(ws, 22, p["COLS"].get("AMOUNT"), round(total, 2))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

charge = st.file_uploader("Upload Charge Sheet", ["xlsx"])
tpl = st.file_uploader("Upload Template", ["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Member Number")
provider = st.text_input("Provider", "CIMAS")

if charge and st.button("Load Charge Sheet"):
    st.session_state.df = load_charge_sheet(charge)

if "df" in st.session_state:
    df = st.session_state.df

    cat = st.selectbox("Category", df["CATEGORY"].unique())
    sub = st.selectbox(
        "Subcategory",
        df[df["CATEGORY"] == cat]["SUBCATEGORY"].dropna().unique()
    )

    scans = df[(df["CATEGORY"] == cat) & (df["SUBCATEGORY"] == sub)].reset_index(drop=True)

    sel = st.multiselect(
        "Select scans",
        scans.index,
        format_func=lambda i: scans.at[i, "SCAN"]
    )

    rows = [scans.iloc[i].to_dict() for i in sel]

    if rows:
        st.dataframe(pd.DataFrame(rows)[["SCAN", "MODIFIER", "AMOUNT"]])

        if tpl and st.button("Generate & Download"):
            out = fill_template(tpl, patient, member, provider, rows)
            st.download_button("Download", out, "quotation.xlsx")
