import streamlit as st
import pandas as pd
import openpyxl
import io
from copy import copy

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config / heuristics ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN",
    "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND"
}
GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# ---------- Helpers ----------
def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def u(x):
    return clean_text(x).upper()

def safe_int(x, default=1):
    try:
        return int(float(str(x).replace(",", "").strip()))
    except:
        return default

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "").strip())
    except:
        return default

# ---------- Parser ----------
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

        exam_u = exam.upper()

        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        if exam_u in GARBAGE_KEYS:
            continue

        tariff_blank = pd.isna(r["B_TARIFF"])
        amt_blank = pd.isna(r["E_AMOUNT"])

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

# ---------- Excel helpers ----------
def write_safe(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    try:
        cell.value = value
    except:
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                ws[mr.coord.split(":")[0]].value = value
                return

def find_template_positions(ws):
    pos = {}
    headers = ["DESCRIPTION", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]

    for row in ws.iter_rows(min_row=1, max_row=300):
        for cell in row:
            if not cell.value:
                continue

            t = u(cell.value)

            if "PATIENT" in t:
                pos["patient"] = (cell.row, cell.column)
            if "MEMBER" in t:
                pos["member"] = (cell.row, cell.column)
            if "PROVIDER" in t or "EXAMINATION" in t:
                pos["provider"] = (cell.row, cell.column)

            if any(h in t for h in headers):
                pos.setdefault("cols", {})
                pos["table_start"] = cell.row + 1

                for h in headers:
                    if h in t:
                        pos["cols"][h] = cell.column

    return pos

def replace_after_colon(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            cell = ws[mr.coord.split(":")[0]]
            break

    old = str(cell.value) if cell.value else ""
    cell.value = f"{old.split(':')[0]}: {value}" if ":" in old else value

def write_preserve_borders(ws, addr, value):
    cell = ws[addr]
    merged = None

    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            merged = str(mr)
            ws.unmerge_cells(merged)
            cell = ws[mr.coord.split(":")[0]]
            break

    border = copy(cell.border)
    font = copy(cell.font)
    fill = copy(cell.fill)
    align = copy(cell.alignment)

    cell.value = value
    cell.border = border
    cell.font = font
    cell.fill = fill
    cell.alignment = align

    if merged:
        ws.merge_cells(merged)

# ---------- Fill Template ----------
def fill_excel_template(template_file, patient, member, provider, rows):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active
    pos = find_template_positions(ws)

    if "patient" in pos:
        replace_after_colon(ws, *pos["patient"], patient)
    if "member" in pos:
        replace_after_colon(ws, *pos["member"], member)
    if "provider" in pos:
        replace_after_colon(ws, *pos["provider"], provider)

    rowptr = pos.get("table_start", 20)

    for r in rows:
        write_safe(ws, rowptr, 1, r["SCAN"])        # Description
        write_safe(ws, rowptr, 2, r["TARIFF"])      # Tariff
        write_safe(ws, rowptr, 3, r["MODIFIER"])    # MOD (FORCED COLUMN C)
        write_safe(ws, rowptr, 4, r["QTY"])
        write_safe(ws, rowptr, 5, r["AMOUNT"])
        rowptr += 1

    total = sum(safe_float(r["AMOUNT"]) for r in rows)
    write_preserve_borders(ws, "G22", total)
    write_preserve_borders(ws, "G41", total)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------- UI ----------
st.title("Medical Quotation Generator")

uploaded_charge = st.file_uploader("Upload Charge Sheet", type="xlsx")
uploaded_template = st.file_uploader("Upload Template", type="xlsx")

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member")
provider = st.text_input("Provider", value="CIMAS")

if uploaded_charge and st.button("Parse Charge Sheet"):
    st.session_state.df = load_charge_sheet(uploaded_charge)

if "df" in st.session_state:
    df = st.session_state.df

    cats = sorted(df["CATEGORY"].dropna().unique())
    main = st.selectbox("Main Category", cats)

    subs = sorted(df[df["CATEGORY"] == main]["SUBCATEGORY"].dropna().unique())
    sub = st.selectbox("Sub Category", subs) if subs else None

    subset = df[(df["CATEGORY"] == main) & (df["SUBCATEGORY"] == sub)] if sub else df[df["CATEGORY"] == main]

    subset = subset.reset_index(drop=True)
    subset["label"] = subset.apply(lambda r: f"{r['SCAN']} | {r['AMOUNT']}", axis=1)

    picks = st.multiselect(
        "Select scans",
        options=list(range(len(subset))),
        format_func=lambda i: subset.at[i, "label"]
    )

    if picks:
        selected = subset.iloc[picks][
            ["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]
        ].copy()

        edited = st.data_editor(
            selected,
            column_config={
                "SCAN": st.column_config.TextColumn("Description"),
                "TARIFF": st.column_config.NumberColumn("Tariff", disabled=True),
                "MODIFIER": st.column_config.TextColumn("MOD", disabled=True),
                "QTY": st.column_config.NumberColumn("Qty", disabled=True),
                "AMOUNT": st.column_config.NumberColumn("Amount", disabled=True),
            },
            use_container_width=True,
            num_rows="fixed"
        )

        rows = edited.to_dict("records")
        total = sum(safe_float(r["AMOUNT"]) for r in rows)
        st.markdown(f"**Total: {total:.2f}**")

        if uploaded_template and st.button("Generate Quotation"):
            out = fill_excel_template(uploaded_template, patient, member, provider, rows)
            st.download_button(
                "Download Excel",
                out,
                f"quotation_{patient or 'patient'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
