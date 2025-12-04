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
GARBAGE_KEYS = {"FF", "TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# ---------- Helpers ----------
def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()

def u(x) -> str:
    return clean_text(x).upper()

def safe_int(x, default=1):
    try:
        v = pd.to_numeric(x, errors="coerce")
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return default
        return int(v)
    except Exception:
        return default

def safe_float(x, default=0.0):
    try:
        v = pd.to_numeric(x, errors="coerce")
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return default
        return float(v)
    except Exception:
        return default

# ---------- Parser ----------
def load_charge_sheet(file) -> pd.DataFrame:
    # read without header (sheet often lacks proper headers)
    df_raw = pd.read_excel(file, header=None, dtype=object)
    # we expect at least 5 columns; pad if necessary
    while df_raw.shape[1] < 5:
        df_raw[df_raw.shape[1]] = None
    df_raw = df_raw.iloc[:, :5]
    df_raw.columns = ["A_EXAM", "B_TARIFF", "C_MOD", "D_QTY", "E_AMOUNT"]

    structured = []
    current_category: Optional[str] = None
    current_subcategory: Optional[str] = None

    # Try to find the header row where the row contains words like TARIFF / AMOUNT
    header_row_idx = None
    for idx, row in df_raw.iterrows():
        row_text = " ".join([u(row[col]) for col in df_raw.columns])
        if "TARRIF" in row_text or "TARIFF" in row_text or "AMOUNT" in row_text:
            header_row_idx = idx
            break
    # If header row found, drop everything above it except categories/subcats
    if header_row_idx is not None:
        # We'll iterate from header_row_idx+1; but also collect categories above header if present
        start_idx = header_row_idx + 1
        # check above for a category-like text (single cell in col A)
        for i in range(0, header_row_idx + 1):
            val = u(df_raw.at[i, "A_EXAM"])
            if val and val in MAIN_CATEGORIES:
                current_category = val
                break
    else:
        start_idx = 0

    # iterate rows and form structured rows
    for i in range(start_idx, len(df_raw)):
        r = df_raw.iloc[i]
        exam = u(r["A_EXAM"])
        tariff_raw = r["B_TARIFF"]
        amt_raw = r["E_AMOUNT"]

        # If this row explicitly names a MAIN CATEGORY (e.g. ULTRA SOUND DOPPLERS)
        if exam in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        # If row has no tariff & no amount and A_EXAM is non-empty -> treat as subcategory
        tariff_is_blank = pd.isna(tariff_raw) or str(tariff_raw).strip() == ""
        amount_is_blank = pd.isna(amt_raw) or str(amt_raw).strip() == ""
        if exam and tariff_is_blank and amount_is_blank:
            # ignore if garbage
            if exam in GARBAGE_KEYS:
                continue
            # treat as subcategory
            current_subcategory = exam
            continue

        # If row is garbage key in A_EXAM, skip
        if exam in GARBAGE_KEYS:
            continue

        # Otherwise, treat as a scan row (even if some fields missing)
        row_tariff = safe_float(tariff_raw, default=None)
        row_mod = clean_text(r["C_MOD"])
        row_qty = safe_int(r["D_QTY"], default=1)
        row_amt = safe_float(amt_raw, default=0.0)

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": clean_text(r["A_EXAM"]),
            "TARIFF": row_tariff,
            "MODIFIER": row_mod,
            "QTY": row_qty,
            "AMOUNT": row_amt
        })

    df_struct = pd.DataFrame(structured, columns=[
        "CATEGORY", "SUBCATEGORY", "SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"
    ])
    # If parser produced no rows, attempt fallback: treat every row with a B_TARIFF as scan
    if df_struct.empty:
        for _, r in df_raw.iterrows():
            if not (pd.isna(r["B_TARIFF"]) and pd.isna(r["E_AMOUNT"])):
                df_struct = df_struct.append({
                    "CATEGORY": None,
                    "SUBCATEGORY": None,
                    "SCAN": clean_text(r["A_EXAM"]),
                    "TARIFF": safe_float(r["B_TARIFF"], default=None),
                    "MODIFIER": clean_text(r["C_MOD"]),
                    "QTY": safe_int(r["D_QTY"], default=1),
                    "AMOUNT": safe_float(r["E_AMOUNT"], default=0.0)
                }, ignore_index=True)
    return df_struct

# ---------- Excel template filler ----------
def write_safe(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    try:
        cell.value = value
    except Exception:
        # merged cell fallback: set top-left of merged range
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                top = mr.coord.split(":")[0]
                ws[top].value = value
                return

def find_template_positions(ws):
    """
    Try to locate:
      - patient cell (cell next to label PATIENT)
      - member cell (cell next to MEMBER)
      - provider cell (cell next to PROVIDER / EXAMINATION)
      - the description header row (DESCRIPTION) and its column
      - a TOTAL label to place the overall total (optional)
    Returns dict with row/col positions.
    """
    pos = {}
    for row in ws.iter_rows(min_row=1, max_row=200, min_col=1, max_col=50):
        for cell in row:
            if cell.value:
                t = u(cell.value)
                if "PATIENT" in t and "patient_cell" not in pos:
                    pos["patient_cell"] = (cell.row, cell.column + 1)
                if "MEMBER" in t and "member_cell" not in pos:
                    pos["member_cell"] = (cell.row, cell.column + 1)
                if ("PROVIDER" in t or "EXAMINATION" in t) and "provider_cell" not in pos:
                    pos["provider_cell"] = (cell.row, cell.column + 1)
                if "DESCRIPTION" in t and "table_start_row" not in pos:
                    pos["table_start_row"] = cell.row + 1
                    pos["desc_col"] = cell.column
                if t.strip() == "TOTAL" and "total_cell" not in pos:
                    pos["total_cell"] = (cell.row, cell.column + 6)
    return pos

def fill_excel_template(template_file, patient, member, provider, scan_rows):
    """
    scan_rows: list/df of scan dicts/rows to add (can be a single row or many)
    """
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    # Fill patient info
    if "patient_cell" in pos:
        r, c = pos["patient_cell"]
        write_safe(ws, r, c, patient)
    if "member_cell" in pos:
        r, c = pos["member_cell"]
        write_safe(ws, r, c, member)
    if "provider_cell" in pos:
        r, c = pos["provider_cell"]
        write_safe(ws, r, c, provider)

    # Fill scans into table_start_row onward
    if "table_start_row" in pos:
        rowptr = pos["table_start_row"]
        desc_col = pos["desc_col"]
        for sr in scan_rows:
            write_safe(ws, rowptr, desc_col, sr.get("SCAN"))
            write_safe(ws, rowptr, desc_col + 1, sr.get("TARIFF"))
            write_safe(ws, rowptr, desc_col + 2, sr.get("MODIFIER"))
            write_safe(ws, rowptr, desc_col + 3, sr.get("QTY"))
            write_safe(ws, rowptr, desc_col + 4, sr.get("AMOUNT"))
            rowptr += 1

        # write total if position exists
        if "total_cell" in pos:
            total = sum([safe_float(s.get("AMOUNT"), 0.0) for s in scan_rows])
            r, c = pos["total_cell"]
            write_safe(ws, r, c, total)
    else:
        # fallback: try to write into first sheet area top-left
        rowptr = 25
        desc_col = 1
        for sr in scan_rows:
            write_safe(ws, rowptr, desc_col, sr.get("SCAN"))
            write_safe(ws, rowptr, desc_col + 1, sr.get("TARIFF"))
            write_safe(ws, rowptr, desc_col + 2, sr.get("MODIFIER"))
            write_safe(ws, rowptr, desc_col + 3, sr.get("QTY"))
            write_safe(ws, rowptr, desc_col + 4, sr.get("AMOUNT"))
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
        st.write("First 50 parsed rows:")
        st.dataframe(df.head(50))

    # categories
    cats = [c for c in sorted(df["CATEGORY"].dropna().unique())] if "CATEGORY" in df.columns else []
    if not cats:
        # fallback: allow selecting by SUBCATEGORY or show all scans
        subs = [s for s in sorted(df["SUBCATEGORY"].dropna().unique())] if "SUBCATEGORY" in df.columns else []
        if subs:
            st.warning("No main categories detected; choose a Subcategory instead.")
            subsel = st.selectbox("Select Subcategory", subs)
            scans_for_sub = df[df["SUBCATEGORY"] == subsel]
        else:
            st.warning("No categories/subcategories detected; showing all scans.")
            scans_for_sub = df
    else:
        main_sel = st.selectbox("Select Main Category", ["-- choose --"] + cats)
        if main_sel == "-- choose --":
            st.info("Please select a main category.")
            st.stop()
        subs = [s for s in sorted(df[df["CATEGORY"] == main_sel]["SUBCATEGORY"].dropna().unique())]
        if not subs:
            st.warning("No subcategories found under this category; showing scans directly.")
            scans_for_sub = df[df["CATEGORY"] == main_sel]
        else:
            subsel = st.selectbox("Select Subcategory", subs)
            scans_for_sub = df[(df["CATEGORY"] == main_sel) & (df["SUBCATEGORY"] == subsel)]

    # show scan list and allow multiple selection
    if scans_for_sub.empty:
        st.warning("No scans available for the current selection.")
    else:
        scans_for_sub = scans_for_sub.reset_index(drop=True)
        # create display labels
        scans_for_sub["label"] = scans_for_sub.apply(
            lambda r: f"{r['SCAN']}  | Tariff: {r['TARIFF']}  | Amt: {r['AMOUNT']}", axis=1
        )
        sel_indices = st.multiselect("Select scans to add to quotation (you can select multiple)", options=list(range(len(scans_for_sub))),
                                     format_func=lambda i: scans_for_sub.at[i, "label"])
        # build selected list
        selected_rows = [scans_for_sub.iloc[i].to_dict() for i in sel_indices]
        if selected_rows:
            st.write("Selected scans:")
            display_df = pd.DataFrame(selected_rows)[["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]]
            st.dataframe(display_df)
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
