import streamlit as st
import pandas as pd
import openpyxl
import io
import math
from typing import Optional

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config / heuristics ----------
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
    except:
        return default

def safe_float(x, default=0.0):
    try:
        x_str = str(x).replace(",", "").strip()
        return float(x_str)
    except:
        return default

# ---------- Parser ----------
def load_charge_sheet(file) -> pd.DataFrame:
    df_raw = pd.read_excel(file, header=None, dtype=object)

    while df_raw.shape[1] < 5:
        df_raw[df_raw.shape[1]] = None
    df_raw = df_raw.iloc[:, :5]
    df_raw.columns = ["A_EXAM", "B_TARIFF", "C_MOD", "D_QTY", "E_AMOUNT"]

    structured = []
    current_category = None
    current_subcategory = None

    for idx, r in df_raw.iterrows():
        exam = clean_text(r["A_EXAM"])
        if exam == "":
            continue

        exam_u = exam.upper()

        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        if exam_u == "FF":
            structured.append({
                "CATEGORY": current_category,
                "SUBCATEGORY": current_subcategory,
                "SCAN": "FF",
                "TARIFF": safe_float(r["B_TARIFF"], None),
                "MODIFIER": "",
                "QTY": safe_int(r["D_QTY"], 1),
                "AMOUNT": safe_float(r["E_AMOUNT"], 0.0)
            })
            continue

        if exam_u in GARBAGE_KEYS:
            continue

        tariff_blank = pd.isna(r["B_TARIFF"]) or str(r["B_TARIFF"]).strip() in ["", "nan", "NaN", "None"]
        amt_blank = pd.isna(r["E_AMOUNT"]) or str(r["E_AMOUNT"]).strip() in ["", "nan", "NaN", "None"]

        if tariff_blank and amt_blank:
            current_subcategory = exam
            continue

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "TARIFF": safe_float(r["B_TARIFF"], None),
            "MODIFIER": clean_text(r["C_MOD"]),
            "QTY": safe_int(r["D_QTY"], 1),
            "AMOUNT": safe_float(r["E_AMOUNT"], 0.0)
        })

    return pd.DataFrame(structured)

# ---------- Excel Template Mapping ----------
def write_safe(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    try:
        cell.value = value
    except:
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                topcell = mr.coord.split(":")[0]
                ws[topcell].value = value
                return

def find_template_positions(ws):
    pos = {}
    for row in ws.iter_rows(min_row=1, max_row=300):
        for cell in row:
            if cell.value:
                t = u(cell.value)
                if "PATIENT" in t and "patient_cell" not in pos:
                    pos["patient_cell"] = (cell.row, cell.column + 1)

                if "MEMBER" in t and "member_cell" not in pos:
                    pos["member_cell"] = (cell.row, cell.column + 1)

                if ("PROVIDER" in t or "EXAMINATION" in t) and "provider_cell" not in pos:
                    pos["provider_cell"] = (cell.row, cell.column + 1)

                headers = ["DESCRIPTION", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]
                if any(h in t for h in headers) and "cols" not in pos:
                    pos["cols"] = {}
                    pos["table_start_row"] = cell.row + 1

                for h in headers:
                    if h in t:
                        pos["cols"][h] = cell.column
    return pos

# ---------- Template Filler (with TOTAL column update) ----------
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
        start_row = pos["table_start_row"]
        cols = pos["cols"]

        rowptr = start_row

        # --- Write Scan Rows ---
        for sr in scan_rows:
            write_safe(ws, rowptr, cols.get("DESCRIPTION"), sr.get("SCAN"))
            write_safe(ws, rowptr, cols.get("TARRIF"), sr.get("TARIFF"))
            write_safe(ws, rowptr, cols.get("MOD"), sr.get("MODIFIER"))
            write_safe(ws, rowptr, cols.get("QTY"), sr.get("QTY"))
            write_safe(ws, rowptr, cols.get("FEES"), sr.get("AMOUNT"))
            rowptr += 1

        # --- WRITE TOTAL IN NEXT COLUMN AFTER AMOUNT (ON FIRST SCAN ROW) ---
        total_amt = sum([safe_float(r["AMOUNT"], 0.0) for r in scan_rows])

        col_amount = cols.get("AMOUNT")
        col_total = col_amount + 1  # NEXT COLUMN (keeps blue diagonal intact)

        write_safe(ws, start_row, col_total, total_amt)

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

    cats = [c for c in sorted(df["CATEGORY"].dropna().unique())]

    if not cats:
        subs = [s for s in sorted(df["SUBCATEGORY"].dropna().unique())]
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

        sel_indices = st.multiselect(
            "Select scans to add to quotation (you can select multiple)",
            options=list(range(len(scans_for_sub))),
            format_func=lambda i: scans_for_sub.at[i, "label"]
        )

        selected_rows = [scans_for_sub.iloc[i].to_dict() for i in sel_indices]

        if selected_rows:
            st.dataframe(pd.DataFrame(selected_rows)[["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]])

            total_amt = sum([safe_float(r["AMOUNT"], 0.0) for r in selected_rows])
            st.markdown(f"**Total Amount:** {total_amt:.2f}")

            if uploaded_template:
                if st.button("Generate Quotation and Download Excel"):
                    out = fill_excel_template(uploaded_template, patient, member, provider, selected_rows)
                    st.download_button(
                        "Download Quotation",
                        data=out,
                        file_name=f"quotation_{patient or 'patient'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("Upload a quotation template to enable download.")
        else:
            st.info("No scans selected yet. Choose scans to add to the quotation.")
else:
    st.info("Upload a charge sheet to begin parsing.")
