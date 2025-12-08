import streamlit as st
import pandas as pd
import openpyxl
import io
from copy import copy
from datetime import datetime

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config / heuristics ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN",
    "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND",
    "MRI"
}
GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", "", "CONSUMABLES"}

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
    current_category = None
    current_subcategory = None

    for idx, r in df_raw.iterrows():
        exam = clean_text(r["A_EXAM"])
        if exam == "":
            continue

        exam_u = exam.upper()

        # Main category
        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        tariff_blank = pd.isna(r["B_TARIFF"]) or str(r["B_TARIFF"]).strip() in ["", "nan", "NaN", "None"]
        amt_blank = pd.isna(r["E_AMOUNT"]) or str(r["E_AMOUNT"]).strip() in ["", "nan", "NaN", "None"]

        # Subcategory header
        if tariff_blank and amt_blank:
            current_subcategory = exam
            continue

        # MRI-specific FF rows: include under subcategory
        if exam_u == "FF" and current_category == "MRI" and current_subcategory is not None:
            structured.append({
                "CATEGORY": current_category,
                "SUBCATEGORY": current_subcategory,
                "SCAN": current_subcategory,  # Use subcategory name
                "TARIFF": safe_float(r["B_TARIFF"], None),
                "MODIFIER": clean_text(r["C_MOD"]),
                "QTY": safe_int(r["D_QTY"], 1),
                "AMOUNT": safe_float(r["E_AMOUNT"], 0.0)
            })
            continue

        # Skip garbage for other categories
        if exam_u in GARBAGE_KEYS:
            continue

        # Real scan rows
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

# ---------- Excel Template Helpers ----------
def write_safe(ws, r, c, value):
    if c is None:
        return
    try:
        cell = ws.cell(row=r, column=c)
    except Exception:
        return
    try:
        cell.value = value
    except Exception:
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
                    pos["patient_cell"] = (cell.row, cell.column)
                if "MEMBER" in t and "member_cell" not in pos:
                    pos["member_cell"] = (cell.row, cell.column)
                if ("PROVIDER" in t or "EXAMINATION" in t) and "provider_cell" not in pos:
                    pos["provider_cell"] = (cell.row, cell.column)
                if "DATE" in t and "date_cell" not in pos:
                    pos["date_cell"] = (cell.row, cell.column)

                headers = ["DESCRIPTION", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]
                if any(h in t for h in headers) and "cols" not in pos:
                    pos["cols"] = {}
                    pos["table_start_row"] = cell.row + 1

                for h in headers:
                    if h in t:
                        pos.setdefault("cols", {})[h] = cell.column
    return pos

def replace_after_colon_in_same_cell(ws, row, col, new_value):
    cell = ws.cell(row=row, column=col)
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            cell = ws[mr.coord.split(":")[0]]
            break
    old = str(cell.value) if cell.value else ""
    if ":" in old:
        left = old.split(":", 1)[0]
        cell.value = f"{left}: {new_value}"
    else:
        cell.value = new_value

def write_value_preserve_borders(ws, cell_address, value):
    cell = ws[cell_address]
    merged_range = None
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            merged_range = mr
            cell = ws[mr.coord.split(":")[0]]
            ws.unmerge_cells(str(mr))
            break
    border = copy(cell.border)
    font = copy(cell.font)
    fill = copy(cell.fill)
    alignment = copy(cell.alignment)

    cell.value = value

    cell.border = border
    cell.font = font
    cell.fill = fill
    cell.alignment = alignment

    if merged_range:
        ws.merge_cells(str(merged_range))

# ---------- Fill Template ----------
def fill_excel_template(template_file, patient, member, provider, scan_rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    if "patient_cell" in pos:
        r, c = pos["patient_cell"]
        replace_after_colon_in_same_cell(ws, r, c, patient)
    if "member_cell" in pos:
        r, c = pos["member_cell"]
        replace_after_colon_in_same_cell(ws, r, c, member)
    if "provider_cell" in pos:
        r, c = pos["provider_cell"]
        replace_after_colon_in_same_cell(ws, r, c, provider)
    if "date_cell" in pos:
        r, c = pos["date_cell"]
        today_str = datetime.now().strftime("%d/%m/%Y")
        ws.cell(row=r+1, column=c, value=today_str)

    if "table_start_row" in pos and "cols" in pos:
        start_row = pos["table_start_row"]
        cols = pos["cols"]
        rowptr = start_row
        for sr in scan_rows:
            write_safe(ws, rowptr, cols.get("DESCRIPTION"), sr.get("SCAN"))
            write_safe(ws, rowptr, cols.get("TARRIF"), sr.get("TARIFF"))
            mod_value = sr.get("MODIFIER")
            if mod_value:
                write_safe(ws, rowptr, cols.get("MOD"), mod_value)
            write_safe(ws, rowptr, cols.get("QTY"), sr.get("QTY"))
            write_safe(ws, rowptr, cols.get("FEES"), sr.get("AMOUNT"))
            rowptr += 1

        total_amt = sum([safe_float(r.get("AMOUNT", 0.0), 0.0) for r in scan_rows])

        try:
            write_value_preserve_borders(ws, "G22", total_amt)
        except Exception:
            try:
                ws["G22"].value = total_amt
            except Exception:
                pass
        try:
            write_value_preserve_borders(ws, "G41", total_amt)
        except Exception:
            try:
                ws["G41"].value = total_amt
            except Exception:
                pass

        for row in ws.iter_rows(min_row=1, max_row=300):
            for cell in row:
                if cell.value and str(cell.value).strip().upper() == "TOTAL":
                    target = f"G{cell.row}"
                    try:
                        write_value_preserve_borders(ws, target, total_amt)
                    except Exception:
                        ws[target].value = total_amt
                    break

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------- Streamlit UI ----------
st.title("Medical Quotation Generator — MRI FF Fix")

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
            st.session_state.selected_rows = []
            st.success("Charge sheet parsed successfully.")
        except Exception as e:
            st.error(f"Failed to parse charge sheet: {e}")
            st.stop()

if "parsed_df" in st.session_state:
    df = st.session_state.parsed_df

    if debug_mode:
        st.write("Parsed DataFrame columns:", df.columns.tolist())
        st.dataframe(df.head(200))

    cats = [c for c in sorted(df["CATEGORY"].dropna().unique())] if "CATEGORY" in df.columns else []

    if not cats:
        st.warning("No categories found in the uploaded charge sheet.")
    else:
        col1, col2 = st.columns([3, 4])
        with col1:
            main_sel = st.selectbox("Select Main Category", ["-- choose --"] + cats)
        if main_sel != "-- choose --":
            subs = [s for s in sorted(df[df["CATEGORY"] == main_sel]["SUBCATEGORY"].dropna().unique())]
            if subs:
                sub_sel = st.selectbox("Select Subcategory", ["-- all --"] + subs)
            else:
                sub_sel = "-- all --"

            if sub_sel == "-- all --":
                scans_for_cat = df[df["CATEGORY"] == main_sel].reset_index(drop=True)
            else:
                scans_for_cat = df[
                    (df["CATEGORY"] == main_sel) &
                    (df["SUBCATEGORY"] == sub_sel)
                ].reset_index(drop=True)

            # Keep only real scans
            scans_filtered = scans_for_cat[
                (~scans_for_cat["SCAN"].isin(GARBAGE_KEYS)) &
                (scans_for_cat["SCAN"].notna())
            ]
            st.session_state.selected_rows = scans_filtered.to_dict(orient="records")

    st.markdown("---")
    st.subheader("Selected Scans")
    if "selected_rows" not in st.session_state or len(st.session_state.selected_rows) == 0:
        st.info("No real scans available in this category/subcategory.")
        selected_df = pd.DataFrame(columns=["SCAN", "TARRIF", "MODIFIER", "QTY", "AMOUNT"])
        st.dataframe(selected_df)
    else:
        sel_df = pd.DataFrame(st.session_state.selected_rows)
        display_df = sel_df[["SCAN", "TARRIF", "MODIFIER", "QTY", "AMOUNT"]]
        st.dataframe(display_df.reset_index(drop=True))

        total_amt = sum([safe_float(r.get("AMOUNT", 0.0), 0.0) for r in st.session_state.selected_rows])
        st.markdown(f"**Total Amount:** {total_amt:.2f}")

        if uploaded_template and st.button("Generate Quotation and Download Excel"):
            try:
                out = fill_excel_template(uploaded_template, patient, member, provider, st.session_state.selected_rows)
                st.download_button(
                    "Download Quotation",
                    data=out,
                    file_name=f"quotation_{(patient or 'patient').replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Failed to generate quotation: {e}")
else:
    st.info("Upload a charge sheet to begin.")
