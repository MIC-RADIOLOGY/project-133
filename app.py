# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime
from copy import copy

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

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
    except Exception:
        return default

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "").strip())
    except Exception:
        return default

# ------------------------------------------------------------
# PARSER
# ------------------------------------------------------------
def load_charge_sheet(file):
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

        exam_u = exam.upper().strip()
        if exam_u in MAIN_CATEGORIES or exam_u.endswith("SCAN") or exam_u in {"XRAY", "MRI", "ULTRASOUND"}:
            MAIN_CATEGORIES.add(exam_u)
            current_category = exam
            current_subcategory = None
            continue

        if exam_u in GARBAGE_KEYS:
            continue

        if clean_text(r["B_TARIFF"]) == "" and clean_text(r["E_AMOUNT"]) == "":
            current_subcategory = exam
            continue

        if not current_category:
            continue

        is_main_scan = exam_u not in COMPONENT_KEYS

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "IS_MAIN_SCAN": is_main_scan,
            "TARIFF": safe_float(r["B_TARIFF"], None),
            "MODIFIER": str(clean_text(r["C_MOD"])),
            "QTY": safe_int(r["D_QTY"], 1),
            "AMOUNT": safe_float(r["E_AMOUNT"], 0.0)
        })

    return pd.DataFrame(structured)

# ------------------------------------------------------------
# EXCEL HELPERS
# ------------------------------------------------------------
def write_to_data_sheet(wb, scan_rows):
    if "Data" in wb.sheetnames:
        ws_data = wb["Data"]
    else:
        ws_data = wb.create_sheet("Data")
        ws_data.sheet_state = 'hidden'

    headers = ["DESCRIPTION", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]
    for col, h in enumerate(headers, 1):
        ws_data.cell(row=1, column=col, value=h)

    for r, row in enumerate(scan_rows, start=2):
        ws_data.cell(row=r, column=1, value=row["SCAN"])
        ws_data.cell(row=r, column=2, value=row["TARIFF"])
        ws_data.cell(row=r, column=3, value=row["MODIFIER"])
        ws_data.cell(row=r, column=4, value=row["QTY"])
        ws_data.cell(row=r, column=5, value=row["AMOUNT"])

# ------------------------------------------------------------
# TEMPLATE FILL
# ------------------------------------------------------------
def fill_excel_template(template_file, patient, member, provider, scan_rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    # Fill patient info
    for row in ws.iter_rows(min_row=1, max_row=50):
        for cell in row:
            if not cell.value:
                continue
            t = str(cell.value).upper()
            if "PATIENT" in t:
                cell.value = f"{cell.value} {patient}"
            elif "MEMBER" in t:
                cell.value = f"{cell.value} {member}"
            elif "PROVIDER" in t or "MEDICAL AID" in t:
                cell.value = f"{cell.value} {provider}"
            elif t.strip() == "DATE":
                ws.cell(row=cell.row + 1, column=cell.column, value=datetime.today().strftime("%d/%m/%Y"))

    # Write data to hidden sheet
    write_to_data_sheet(wb, scan_rows)

    # At this point, the template should have formulas pointing to Data sheet:
    # Example:
    # = 'Data'!A2 for DESCRIPTION
    # =SUM('Data'!E2:E50) for TOTAL
    # So we don’t overwrite template cells directly

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ------------------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if uploaded_charge and st.button("Load & Parse Charge Sheet"):
    st.session_state.df = load_charge_sheet(uploaded_charge)
    st.success("Charge sheet parsed successfully.")

if "df" in st.session_state:
    df = st.session_state.df

    main_sel = st.selectbox("Select Main Category", sorted(df["CATEGORY"].dropna().unique()))
    subcats = sorted(df[df["CATEGORY"] == main_sel]["SUBCATEGORY"].dropna().unique())
    sub_sel = st.selectbox("Select Subcategory", subcats) if subcats else None

    scans = (df[(df["CATEGORY"] == main_sel) & (df["SUBCATEGORY"] == sub_sel)]
             if sub_sel else df[df["CATEGORY"] == main_sel]).reset_index(drop=True)

    scans["label"] = scans.apply(lambda r: f"{r['SCAN']} | Tariff {r['TARIFF']} | Amount {r['AMOUNT']}", axis=1)

    selected = st.multiselect("Select scans to include", options=list(range(len(scans))),
                              format_func=lambda i: scans.at[i, "label"])

    selected_rows = [scans.iloc[i].to_dict() for i in selected]

    if selected_rows:
        st.subheader("Edit final description for Excel")
        for i, row in enumerate(selected_rows):
            new_desc = st.text_input(f"Description for '{row['SCAN']}'", value=row['SCAN'], key=f"desc_{i}")
            selected_rows[i]['SCAN'] = new_desc

        st.subheader("Preview of selected scans")
        st.dataframe(pd.DataFrame(selected_rows)[["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]])

        if uploaded_template and st.button("Generate & Download Quotation"):
            safe_name = "".join(c for c in (patient or "patient") if c.isalnum() or c in (" ", "_")).strip()
            out = fill_excel_template(uploaded_template, patient, member, provider, selected_rows)
            st.download_button("Download Quotation", data=out,
                               file_name=f"quotation_{safe_name}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
