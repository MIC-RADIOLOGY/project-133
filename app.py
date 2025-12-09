import streamlit as st
import pandas as pd
import io
from datetime import datetime
import xlsxwriter

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config ----------
MAIN_CATEGORIES = {
    "ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN",
    "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND",
    "MRI"
}
GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

# ---------- Helpers ----------
def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def safe_int(x, default=1):
    try:
        x_str = str(x).replace(",", "").strip()
        return int(float(x_str))
    except Exception:
        return default

def safe_float(x, default=0.0):
    try:
        x_str = str(x).replace(",", "").replace("$", "").strip()
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

        if exam_u in MAIN_CATEGORIES:
            current_category = exam
            current_subcategory = None
            continue

        tariff_blank = pd.isna(r["B_TARIFF"]) or str(r["B_TARIFF"]).strip() in ["", "nan", "NaN", "None"]
        amt_blank = pd.isna(r["E_AMOUNT"]) or str(r["E_AMOUNT"]).strip() in ["", "nan", "NaN", "None"]
        if tariff_blank and amt_blank:
            current_subcategory = exam
            continue

        if exam_u in GARBAGE_KEYS:
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

# ---------- Generate Excel ----------
def generate_quotation_xlsx(patient, member, provider, scan_rows):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Quotation")

    # ---------- Formats ----------
    header_fmt = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border':1, 'align':'center'})
    money_fmt = wb.add_format({'num_format': '$#,##0.00', 'border':1})
    text_fmt = wb.add_format({'border':1})
    blue_total_fmt = wb.add_format({'bg_color': '#00B0F0', 'bold': True, 'border':1, 'align':'center'})

    # ---------- Column widths ----------
    col_widths = [30, 15, 12, 8, 15]
    for col, width in enumerate(col_widths):
        ws.set_column(col, col, width)

    # ---------- Row heights ----------
    ws.set_row(0, 18)
    ws.set_row(1, 18)
    ws.set_row(2, 18)
    ws.set_row(3, 18)
    ws.set_row(6, 20)  # header row
    for r in range(7, 7 + len(scan_rows) + 2):
        ws.set_row(r, 18)

    # ---------- Patient info ----------
    ws.write('A1', 'Patient:')
    ws.write('B1', patient)
    ws.write('A2', 'Member:')
    ws.write('B2', member)
    ws.write('A3', 'Provider:')
    ws.write('B3', provider)
    ws.write('A4', 'Date:')
    ws.write('B4', datetime.now().strftime("%d/%m/%Y"))

    # ---------- Header ----------
    headers = ["Description", "Tariff", "Modifier", "Qty", "Fees"]
    for col, h in enumerate(headers):
        ws.write(6, col, h, header_fmt)

    # ---------- Scan rows ----------
    start_row = 7
    for i, sr in enumerate(scan_rows):
        r = start_row + i
        ws.write(r, 0, sr['SCAN'], text_fmt)
        ws.write(r, 1, sr['TARIFF'], money_fmt)
        ws.write(r, 2, sr['MODIFIER'], text_fmt)
        ws.write(r, 3, sr['QTY'], text_fmt)
        ws.write(r, 4, sr['AMOUNT'], money_fmt)

    # ---------- Blue total line ----------
    total_row = start_row + len(scan_rows)
    ws.merge_range(total_row, 0, total_row, 3, "Total", blue_total_fmt)
    ws.write_formula(total_row, 4, f"=SUM(E{start_row+1}:E{total_row})", blue_total_fmt)

    wb.close()
    output.seek(0)
    return output

# ---------- Streamlit UI ----------
st.title("Medical Quotation Generator — Sample Layout")

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if uploaded_charge:
    if st.button("Load & Parse Charge Sheet"):
        try:
            parsed = load_charge_sheet(uploaded_charge)
            st.session_state.parsed_df = parsed
            st.session_state.selected_rows = parsed.to_dict(orient="records")
            st.success("Charge sheet parsed successfully.")
        except Exception as e:
            st.error(f"Failed to parse charge sheet: {e}")
            st.stop()

if "parsed_df" in st.session_state:
    df = st.session_state.parsed_df
    st.subheader("Selected Scans")
    st.dataframe(df[["SCAN","TARIFF","MODIFIER","QTY","AMOUNT"]].reset_index(drop=True))

    total_amt = sum([safe_float(r.get("AMOUNT", 0.0)) for r in st.session_state.selected_rows])
    st.markdown(f"**Total Amount (Preview):** {total_amt:.2f}")

    if st.button("Generate Quotation and Download Excel"):
        out = generate_quotation_xlsx(patient, member, provider, st.session_state.selected_rows)
        st.download_button(
            "Download Quotation",
            data=out,
            file_name=f"quotation_{(patient or 'patient').replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload a charge sheet to begin.")
