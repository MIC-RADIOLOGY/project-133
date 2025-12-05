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
# keep FF (films) not as garbage
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
        x_str = str(x).replace(",", "").strip().replace("$", "").replace(" ", "")
        return float(x_str)
    except Exception:
        return default

# ---------- Parser ----------
def load_charge_sheet(file) -> pd.DataFrame:
    """
    Read an uploaded charge sheet (file-like or path) and return structured DataFrame with
    columns: CATEGORY, SUBCATEGORY, SCAN, TARIFF, MODIFIER, QTY, AMOUNT
    The parser is robust: it treats 'FF' specially (films), ignores garbage rows, and
    will not auto-add stray tariff rows unless they have a SCAN/description or are FF.
    """
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
        exam_raw = r["A_EXAM"]
        exam = clean_text(exam_raw)
        exam_u = exam.upper()

        # Skip empty A_EXAM unless it's an 'FF' row we want to capture (rare)
        if exam == "" and exam_u != "FF":
            continue

        # MAIN CATEGORY detection
        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        # Explicit FF handling (films) — keep these even if other columns blank
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

        # Skip garbage rows
        if exam_u in GARBAGE_KEYS:
            continue

        # If A_EXAM present but both tariff & amount blank -> subcategory
        tariff_str = "" if pd.isna(r["B_TARIFF"]) else str(r["B_TARIFF"]).strip()
        amount_str = "" if pd.isna(r["E_AMOUNT"]) else str(r["E_AMOUNT"]).strip()
        if tariff_str in ["", "nan", "None", "NaN"] and amount_str in ["", "nan", "None", "NaN"]:
            current_subcategory = exam
            continue

        # Otherwise treat as scan row
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
def write_safe_cell(ws, r, c, value, append=False):
    """
    Write into worksheet cell (row r, col c). r,c are 1-based.
    If append=True and the cell already has a value, append with a space.
    Handles merged cells by writing to top-left cell of the merged range.
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
                top_cell = ws[top]
                if append and top_cell.value:
                    top_cell.value = f"{top_cell.value} {value}"
                else:
                    top_cell.value = value
                return

def find_template_positions(ws):
    """
    Scan worksheet to find:
      - patient_cell: tuple(row,col) of the cell that contains 'FOR PATIENT' or 'PATIENT'
      - member_cell: cell that contains 'MEMBER'
      - provider_cell: cell that contains 'PROVIDER' or 'EXAMINATION'
      - cols mapping for DESCRIPTION, TARRIF, MOD, QTY, FEES, AMOUNT and table_start_row
    """
    pos = {}
    headers = ["DESCRIPTION", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT", "FEE"]
    for row in ws.iter_rows(min_row=1, max_row=400):
        for cell in row:
            if not cell.value:
                continue
            t = u(cell.value)
            # patient / member / provider detections (we will replace text after colon)
            if ("FOR PATIENT" in t or t.strip().startswith("FOR PATIENT")) and "patient_cell" not in pos:
                pos["patient_cell"] = (cell.row, cell.column)
            elif "PATIENT" == t.strip() and "patient_cell" not in pos:
                pos["patient_cell"] = (cell.row, cell.column)

            if "MEMBER" in t and "member_cell" not in pos:
                pos["member_cell"] = (cell.row, cell.column)
            if ("PROVIDER" in t or "EXAMINATION" in t) and "provider_cell" not in pos:
                pos["provider_cell"] = (cell.row, cell.column)

            # detect table headers
            if any(h in t for h in headers):
                if "cols" not in pos:
                    pos["cols"] = {}
                    pos["table_start_row"] = cell.row + 1
                for h in headers:
                    if h in t:
                        # normalize FEES/FEE
                        key = "FEES" if "FEE" in h else h
                        pos["cols"][key] = cell.column
    return pos

def replace_header_field(ws, cell_pos, label_keyword, new_value):
    """
    Replace everything after the colon in a header cell that matches label_keyword.
    Example: cell contains "FOR PATIENT: Old Name" -> becomes "FOR PATIENT: New Name"
    If there is no colon, it will replace the full cell with "LABEL: value".
    cell_pos: (row,col)
    label_keyword: uppercase label to set, e.g. "FOR PATIENT" or "MEMBER NUMBER"
    """
    r, c = cell_pos
    cell = ws.cell(row=r, column=c)
    current = ""
    try:
        current = "" if cell.value is None else str(cell.value)
    except Exception:
        current = ""
    cur_u = u(current)
    # find the colon; if present preserve left part up to colon
    if ":" in current:
        left = current.split(":", 1)[0].strip()
        new_text = f"{left}: {new_value}"
    else:
        # if no colon, use provided label keyword
        new_text = f"{label_keyword}: {new_value}"
    # write into top-left if merged
    try:
        cell.value = new_text
    except Exception:
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                top = mr.coord.split(":")[0]
                ws[top].value = new_text
                return

def fill_template_from_bytes(template_bytes: bytes, patient: str, member: str, provider: str, scan_rows: list):
    """
    template_bytes: bytes from uploaded template file
    scan_rows: list of dicts with keys: SCAN, TARIFF, MODIFIER, QTY, AMOUNT
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    pos = find_template_positions(ws)

    # Replace header fields (overwrite previous name/member)
    if "patient_cell" in pos:
        replace_header_field(ws, pos["patient_cell"], "FOR PATIENT", patient)
    if "member_cell" in pos:
        replace_header_field(ws, pos["member_cell"], "MEMBER NUMBER", member)
    if "provider_cell" in pos:
        replace_header_field(ws, pos["provider_cell"], "PROVIDER", provider)

    # Fill table rows only with selected scan_rows
    if "table_start_row" in pos and "cols" in pos:
        rowptr = pos["table_start_row"]
        cols = pos["cols"]
        for sr in scan_rows:
            # DESCRIPTION
            desc_col = cols.get("DESCRIPTION")
            if desc_col:
                write_safe_cell(ws, rowptr, desc_col, sr.get("SCAN") or "", append=False)
            # TARRIF
            tcol = cols.get("TARRIF")
            if tcol:
                write_safe_cell(ws, rowptr, tcol, sr.get("TARIFF") if sr.get("TARIFF") is not None else "", append=False)
            # MOD
            mcol = cols.get("MOD")
            if mcol:
                write_safe_cell(ws, rowptr, mcol, sr.get("MODIFIER") or "", append=False)
            # QTY
            qcol = cols.get("QTY")
            if qcol:
                write_safe_cell(ws, rowptr, qcol, sr.get("QTY") if sr.get("QTY") is not None else "", append=False)
            # FEES
            fcol = cols.get("FEES")
            if fcol:
                # write per-unit fee
                write_safe_cell(ws, rowptr, fcol, sr.get("AMOUNT") or "", append=False)
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

# Parse charge sheet when user clicks button
if uploaded_charge and st.button("Load & Parse Charge Sheet"):
    try:
        parsed_df = load_charge_sheet(uploaded_charge)
        st.session_state.parsed_df = parsed_df
        st.success("Charge sheet parsed.")
    except Exception as e:
        st.error(f"Failed to parse charge sheet: {e}")
        st.stop()

if "parsed_df" in st.session_state:
    df = st.session_state.parsed_df

    if debug_mode:
        st.write("Parsed DataFrame columns:", df.columns.tolist())
        st.dataframe(df.head(100))

    # categories/subcategories
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

    # show scans and let user select
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
            # compute total using per-unit AMOUNT * QTY
            total_amt = sum([safe_float(r["AMOUNT"], 0.0) * safe_int(r.get("QTY", 1), 1) for r in selected_rows])
            st.markdown(f"**Total Amount:** {total_amt:.2f}")

            # Generate template if template uploaded
            if uploaded_template:
                if st.button("Generate Quotation and Download Excel"):
                    try:
                        # Prepare scan_rows ensuring AMOUNT is per-unit fee
                        scan_rows_for_template = []
                        for r in selected_rows:
                            # If charge sheet AMOUNT appears to be total for that line and QTY>1,
                            # you can change below to per_unit = r['AMOUNT']/r['QTY']. Currently using AMOUNT as per-unit.
                            scan_rows_for_template.append({
                                "SCAN": r.get("SCAN"),
                                "TARIFF": int(r["TARIFF"]) if r.get("TARIFF") is not None else "",
                                "MODIFIER": r.get("MODIFIER", ""),
                                "QTY": safe_int(r.get("QTY"), 1),
                                "AMOUNT": safe_float(r.get("AMOUNT"), 0.0)
                            })

                        template_bytes = uploaded_template.read()
                        out_buf = fill_template_from_bytes(template_bytes, patient, member, provider, scan_rows_for_template)

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
