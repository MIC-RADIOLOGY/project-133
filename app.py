# app.py
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
# keep "FF" out of garbage so films are picked up
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

# ---------- Parser (robust against phantom tariff rows) ----------
def load_charge_sheet(file) -> pd.DataFrame:
    # file: uploaded file-like object or path
    df_raw = pd.read_excel(file, header=None, dtype=object)

    # ensure at least 5 columns
    while df_raw.shape[1] < 5:
        df_raw[df_raw.shape[1]] = None
    df_raw = df_raw.iloc[:, :5]
    df_raw.columns = ["A_EXAM", "B_TARIFF", "C_MOD", "D_QTY", "E_AMOUNT"]

    structured = []
    current_category: Optional[str] = None
    current_subcategory: Optional[str] = None

    for _, r in df_raw.iterrows():
        # robust cleaning
        exam_raw = r["A_EXAM"]
        exam = clean_text(exam_raw)
        exam_u = exam.upper()

        # skip completely blank A_EXAM unless it's FF (we want FF even if label cell blank)
        if exam == "" and exam_u != "FF":
            continue

        # MAIN CATEGORY rows
        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        # explicit FF row handling (films)
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

        # skip garbage keys
        if exam_u in GARBAGE_KEYS:
            continue

        # treat rows with non-empty A_EXAM but empty tariff & amount as subcategory
        tariff_str = "" if pd.isna(r["B_TARIFF"]) else str(r["B_TARIFF"]).strip()
        amount_str = "" if pd.isna(r["E_AMOUNT"]) else str(r["E_AMOUNT"]).strip()
        if tariff_str in ["", "nan", "None", "NaN"] and amount_str in ["", "nan", "None", "NaN"]:
            current_subcategory = exam
            continue

        # otherwise a scan row
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

    df_struct = pd.DataFrame(structured, columns=[
        "CATEGORY", "SUBCATEGORY", "SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"
    ])
    return df_struct

# ---------- Excel template helpers ----------
def write_safe(ws, r, c, value, append=True):
    """
    Write value into worksheet cell (r,c).
    If append=True and there's existing text, append with a space.
    Handles merged cells by writing into top-left of merged range.
    r,c are 1-based integers.
    """
    cell = ws.cell(row=r, column=c)
    try:
        if append and cell.value:
            cell.value = f"{cell.value} {value}"
        else:
            cell.value = value
    except Exception:
        # merged cell fallback
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                top = mr.coord.split(":")[0]
                ws[top].value = f"{ws[top].value or ''} {value}" if append and ws[top].value else value
                return

def find_template_positions(ws):
    """
    Find key positions in the template:
     - patient_cell: cell that contains the label 'FOR PATIENT' (return its (row,col))
     - member_cell: cell containing 'MEMBER' label (return its (row,col))
     - provider_cell: cell containing 'PROVIDER' or 'EXAMINATION' label (return its (row,col))
     - table_start_row and cols mapping for DESCRIPTION,TARRIF,MOD,QTY,FEES,AMOUNT
    """
    pos = {}
    header_keywords = ["DESCRIPTION","TARRIF","MOD","QTY","FEES","AMOUNT"]
    for row in ws.iter_rows(min_row=1, max_row=300):
        for cell in row:
            if not cell.value:
                continue
            t = u(cell.value)
            # For patient/member/provider, keep the same cell (we will append into it)
            if "FOR PATIENT" in t or t.strip().startswith("FOR PATIENT") or "PATIENT" == t.strip():
                if "patient_cell" not in pos:
                    pos["patient_cell"] = (cell.row, cell.column)
            if "MEMBER" in t and "member_cell" not in pos:
                pos["member_cell"] = (cell.row, cell.column)
            if ("PROVIDER" in t or "EXAMINATION" in t) and "provider_cell" not in pos:
                pos["provider_cell"] = (cell.row, cell.column)

            # detect table header row and columns
            if any(h in t for h in header_keywords):
                if "cols" not in pos:
                    pos["cols"] = {}
                    pos["table_start_row"] = cell.row + 1
                for h in header_keywords:
                    if h in t:
                        pos["cols"][h] = cell.column
    return pos

def fill_excel_template_from_bytes(template_bytes: bytes, patient: str, member: str, provider: str, scan_rows: list):
    """
    template_bytes: bytes of the uploaded template
    scan_rows: list of dicts with keys: SCAN, TARIFF, MODIFIER, QTY, AMOUNT
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    pos = find_template_positions(ws)

    # Patient / Member / Provider: append into the label cell (so label stays and value follows)
    if "patient_cell" in pos:
        r, c = pos["patient_cell"]
        write_safe(ws, r, c, patient, append=True)
    if "member_cell" in pos:
        r, c = pos["member_cell"]
        write_safe(ws, r, c, member, append=True)
    if "provider_cell" in pos:
        r, c = pos["provider_cell"]
        write_safe(ws, r, c, provider, append=True)

    # Table: write only the selected scan_rows into consecutive rows starting at table_start_row
    if "table_start_row" in pos and "cols" in pos:
        rowptr = pos["table_start_row"]
        cols = pos["cols"]
        for sr in scan_rows:
            # DESCRIPTION
            desc_col = cols.get("DESCRIPTION")
            if desc_col:
                write_safe(ws, rowptr, desc_col, sr.get("SCAN") or "", append=False)
            # TARRIF
            tcol = cols.get("TARRIF")
            if tcol:
                write_safe(ws, rowptr, tcol, sr.get("TARIFF") if sr.get("TARIFF") is not None else "", append=False)
            # MOD
            mcol = cols.get("MOD")
            if mcol:
                write_safe(ws, rowptr, mcol, sr.get("MODIFIER") or "", append=False)
            # QTY
            qcol = cols.get("QTY")
            if qcol:
                write_safe(ws, rowptr, qcol, sr.get("QTY") if sr.get("QTY") is not None else "", append=False)
            # FEES (amount per unit)
            fcol = cols.get("FEES")
            if fcol:
                # Use AMOUNT as fee-per-unit if that is what's expected, else calculate
                write_safe(ws, rowptr, fcol, sr.get("AMOUNT") or "", append=False)
            rowptr += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------- Streamlit UI ----------
st.title("📄 Medical Quotation Generator (Manual upload)")

debug_mode = st.checkbox("Show parsing debug output", value=False)

st.markdown("**Step 1.** Upload the charge sheet (Excel).")
uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
st.markdown("**Step 2.** Upload the quotation TEMPLATE (Excel).")
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

# Load & parse charge sheet when user clicks button (so they can edit uploads first)
if uploaded_charge and st.button("Load & Parse Charge Sheet"):
    try:
        # read uploaded file into pandas
        parsed_df = load_charge_sheet(uploaded_charge)
        st.session_state.parsed_df = parsed_df
        st.success("Charge sheet parsed.")
    except Exception as e:
        st.error(f"Failed to parse charge sheet: {e}")
        st.stop()

# If parsed, show selection UI
if "parsed_df" in st.session_state:
    df = st.session_state.parsed_df

    if debug_mode:
        st.write("Parsed DataFrame columns:", df.columns.tolist())
        st.dataframe(df.head(80))

    # Build category/subcategory selectors
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

    # Show scans and allow selection
    if scans_for_sub.empty:
        st.warning("No scans available for the current selection.")
    else:
        scans_for_sub = scans_for_sub.reset_index(drop=True)
        scans_for_sub["label"] = scans_for_sub.apply(
            lambda r: f"{r['SCAN']}  | Tariff: {r['TARIFF']}  | Qty: {r['QTY']}  | Amt: {r['AMOUNT']}", axis=1
        )

        sel_indices = st.multiselect(
            "Select scans to add to quotation (you can select multiple)",
            options=list(range(len(scans_for_sub))),
            format_func=lambda i: scans_for_sub.at[i, "label"]
        )

        selected_rows = [scans_for_sub.iloc[i].to_dict() for i in sel_indices]

        if selected_rows:
            st.dataframe(pd.DataFrame(selected_rows)[["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]])
            total_amt = sum([safe_float(r["AMOUNT"], 0.0) * safe_int(r.get("QTY", 1), 1) for r in selected_rows])
            st.markdown(f"**Total Amount:** {total_amt:.2f}")

            # Generate only if template present
            if uploaded_template:
                if st.button("Generate Quotation and Download Excel"):
                    try:
                        # prepare scan_rows for writing: ensure AMOUNT is per-unit fee
                        scan_rows_for_template = []
                        for r in selected_rows:
                            scan_rows_for_template.append({
                                "SCAN": r.get("SCAN"),
                                "TARIFF": int(r["TARIFF"]) if r.get("TARIFF") is not None else "",
                                "MODIFIER": r.get("MODIFIER", ""),
                                "QTY": safe_int(r.get("QTY"), 1),
                                # Use AMOUNT as fee-per-unit — if your charge sheet stores total, you may need to divide by QTY
                                "AMOUNT": safe_float(r.get("AMOUNT"), 0.0)
                            })

                        # read uploaded template bytes
                        template_bytes = uploaded_template.read()
                        out_buf = fill_excel_template_from_bytes(template_bytes, patient, member, provider, scan_rows_for_template)

                        st.download_button(
                            "Download Quotation (Excel)",
                            data=out_buf,
                            file_name=f"quotation_{patient or 'patient'}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success("Quotation generated.")
                    except Exception as e:
                        st.error(f"Failed to generate quotation: {e}")
            else:
                st.info("Upload a quotation template to enable download.")
        else:
            st.info("No scans selected yet. Choose scans to add to the quotation.")
else:
    st.info("Upload a charge sheet and click 'Load & Parse Charge Sheet' to begin.")
