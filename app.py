# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
from typing import Optional

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN", "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND"
}
GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# ---------- Helpers ----------
def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def u(x) -> str:
    return clean_text(x).upper()

def safe_int(x, default=1):
    try:
        x_str = str(x).replace(",", "").strip()
        return int(float(x_str))
    except Exception:
        return default

def safe_float(x, default=0.0):
    try:
        x_str = str(x).replace(",", "").strip()
        return float(x_str)
    except Exception:
        return default

# ---------- Parser ----------
def load_charge_sheet(file) -> pd.DataFrame:
    df_raw = pd.read_excel(file, header=None, dtype=object)
    while df_raw.shape[1] < 5:
        df_raw[df_raw.shape[1]] = None
    df_raw = df_raw.iloc[:, :5]
    df_raw.columns = ["A_EXAM", "B_TARIFF", "C_MOD", "D_QTY", "E_AMOUNT"]

    structured = []
    current_category: Optional[str] = None
    current_subcategory: Optional[str] = None

    for _, r in df_raw.iterrows():
        exam = clean_text(r["A_EXAM"])
        exam_u = exam.upper()

        # Skip empty rows
        if exam == "":
            continue

        # MAIN CATEGORY
        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        # Special case: FF
        if exam_u == "FF":
            row_tariff = safe_float(r["B_TARIFF"], default=None)
            row_amt = safe_float(r["E_AMOUNT"], default=0.0)
            row_qty = safe_int(r["D_QTY"], default=1)
            structured.append({
                "CATEGORY": current_category,
                "SUBCATEGORY": current_subcategory,
                "SCAN": "FF",
                "TARIFF": row_tariff,
                "MODIFIER": "",
                "QTY": row_qty,
                "AMOUNT": row_amt
            })
            continue

        # Skip garbage keys
        if exam_u in GARBAGE_KEYS:
            continue

        # Skip numeric-only DESCRIPTION rows (phantom)
        if exam.replace(".", "").isdigit():
            continue

        # Subcategory row (tariff & amount blank)
        tariff_str = str(r["B_TARIFF"]).strip() if not pd.isna(r["B_TARIFF"]) else ""
        amount_str = str(r["E_AMOUNT"]).strip() if not pd.isna(r["E_AMOUNT"]) else ""
        if tariff_str in ["", "nan", "None", "NaN"] and amount_str in ["", "nan", "None", "NaN"]:
            current_subcategory = exam
            continue

        # Valid scan row
        row_tariff = safe_float(r["B_TARIFF"], default=None)
        row_amt = safe_float(r["E_AMOUNT"], default=0.0)
        row_qty = safe_int(r["D_QTY"], default=1)
        row_mod = clean_text(r["C_MOD"])

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "TARIFF": row_tariff,
            "MODIFIER": row_mod,
            "QTY": row_qty,
            "AMOUNT": row_amt
        })

    return pd.DataFrame(structured)

# ---------- Excel template filler ----------
def write_safe(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    try:
        cell.value = value
    except Exception:
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                top = mr.coord.split(":")[0]
                ws[top].value = value
                return

def find_template_positions(ws):
    pos = {}
    for row in ws.iter_rows(min_row=1, max_row=200):
        for cell in row:
            if cell.value:
                t = u(cell.value)
                if "PATIENT" in t and "patient_cell" not in pos:
                    pos["patient_cell"] = (cell.row, cell.column + 1)
                if "MEMBER" in t and "member_cell" not in pos:
                    pos["member_cell"] = (cell.row, cell.column + 1)
                if ("PROVIDER" in t or "EXAMINATION" in t) and "provider_cell" not in pos:
                    pos["provider_cell"] = (cell.row, cell.column + 1)

                headers = ["DESCRIPTION","TARRIF","MOD","QTY","FEES","AMOUNT"]
                if any(h in t for h in headers) and "cols" not in pos:
                    pos["cols"] = {}
                    pos["table_start_row"] = cell.row + 1
                for h in headers:
                    if h in t:
                        pos["cols"][h] = cell.column
    return pos

def fill_excel_template(template_file, patient, member, provider, scan_rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    if "patient_cell" in pos:
        r, c = pos["patient_cell"]
        write_safe(ws, r, c, patient)
    if "member_cell" in pos:
        r, c = pos["member_cell"]
        write_safe(ws, r, c, member)
    if "provider_cell" in pos:
        r, c = pos["provider_cell"]
        write_safe(ws, r, c, provider)

    if "table_start_row" in pos and "cols" in pos:
        rowptr = pos["table_start_row"]
        cols = pos["cols"]

        for sr in scan_rows:
            write_safe(ws, rowptr, cols.get("DESCRIPTION"), sr.get("SCAN"))
            write_safe(ws, rowptr, cols.get("TARRIF"), sr.get("TARIFF"))
            write_safe(ws, rowptr, cols.get("MOD"), sr.get("MODIFIER"))
            write_safe(ws, rowptr, cols.get("QTY"), sr.get("QTY"))
            write_safe(ws, rowptr, cols.get("FEES"), sr.get("AMOUNT"))
            rowptr += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------- Streamlit UI ----------
st.title("📄 Medical Quotation Generator (Final)")

debug_mode = st.checkbox("Show parsing debug output", value=False)

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if uploaded_charge:
    if st.button("Load & Parse Charge Sheet"):
        try:
            parsed = load_charge_sheet(uploaded_charge)
            st.session_state.parsed_df = parsed
            st.success("Charge sheet parsed.")
        except Exception as e:
            st.error(f"Failed to parse charge sheet: {e}")
            st.stop()

if "parsed_df" in st.session_state:
    df = st.session_state.parsed_df

    if debug_mode:
        st.write("Parsed DataFrame columns:", df.columns.tolist())
        st.dataframe(df.head(50))

    cats = [c for c in sorted(df["CATEGORY"].dropna().unique())] if "CATEGORY" in df.columns else []
    if not cats:
        subs = [s for s in sorted(df["SUBCATEGORY"].dropna().unique())] if "SUBCATEGORY" in df.columns else []
        if subs:
            st.warning("No main categories detected; choose a Subcategory instead.")
            subsel = st.selectbox("Select Subcategory", subs)
            scans_for_sub = df[df["SUBCATEGORY"] == subsel]
        else:
            scans_for_sub = df
    else:
        main_sel = st.selectbox("Select Main Category", ["-- choose --"] + cats)
        if main_sel == "-- choose --":
            st.info("Please select a main category.")
            st.stop()
        subs = [s for s in sorted(df[df["CATEGORY"] == main_sel]["SUBCATEGORY"].dropna().unique())]
        if not subs:
            scans_for_sub = df[df["CATEGORY"] == main_sel]
        else:
            subsel = st.selectbox("Select Subcategory", subs)
            scans_for_sub = df[(df["CATEGORY"] == main_sel) & (df["SUBCATEGORY"] == subsel)]

    if scans_for_sub.empty:
        st.warning("No scans available for the current selection.")
    else:
        scans_for_sub = scans_for_sub.reset_index(drop=True)
        scans_for_sub["label"] = scans_for_sub.apply(
            lambda r: f"{r['SCAN']}  | Tariff: {r['TARIFF']}  | Amt: {r['AMOUNT']}", axis=1
        )
        sel_indices = st.multiselect("Select scans to add to quotation (you can select multiple)",
                                     options=list(range(len(scans_for_sub))),
                                     format_func=lambda i: scans_for_sub.at[i, "label"])
        selected_rows = [scans_for_sub.iloc[i].to_dict() for i in sel_indices]
        if selected_rows:
            st.dataframe(pd.DataFrame(selected_rows)[["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]])
            total_amt = sum([safe_float(r["AMOUNT"], 0.0) for r in selected_rows])
            st.markdown(f"**Total Amount:** {total_amt:.2f}")

            if uploaded_template:
                if st.button("Generate Quotation and Download Excel"):
                    out = fill_excel_template(uploaded_template, patient, member, provider, selected_rows)
                    st.download_button("Download Quotation", data=out,
                                       file_name=f"quotation_{patient or 'patient'}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Upload a quotation template to enable download.")
        else:
            st.info("No scans selected yet. Choose scans to add to the quotation.")
else:
    st.info("Upload a charge sheet to begin parsing.")
