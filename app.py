import streamlit as st
import pandas as pd
import openpyxl
import io
import math
from copy import copy
from openpyxl.styles import Border, Side

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config / heuristics ----------
MAIN_CATEGORIES = {"ULTRA SOUND DOPPLERS", "ULTRA SOUND", "CT SCAN", "FLUROSCOPY", "X-RAY", "XRAY", "ULTRASOUND", "MRI"}
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

# ---------- Load Charge Sheet ----------
def load_charge_sheet(file) -> dict:
    xls = pd.ExcelFile(file)
    df_all = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(file, sheet_name=sheet)
        df_all[sheet.upper()] = df
    return df_all

# ---------- Extract Categories and Tariffs ----------
def extract_category_tariffs(df_all):
    categories = {}
    for sheet, df in df_all.items():
        if "CIMAS USD" not in df.columns:
            continue
        category = None
        for i, row in df.iterrows():
            if isinstance(row["Unnamed: 0"], str) and row["Unnamed: 0"].strip() != "" and (pd.isna(row["CIMAS USD"]) or row["CIMAS USD"]==0):
                category = row["Unnamed: 0"].strip().upper()
                categories[category] = []
            elif category and not pd.isna(row.get("CIMAS USD")) and row.get("Unnamed: 1"):
                categories[category].append((row["Unnamed: 1"], row["CIMAS USD"]))
    return categories

# ---------- Add Selected Tariffs ----------
def add_selected_tariffs(selected_category, selected_tariffs, category_tariffs):
    added_items = []
    lookup = {name: price for name, price in category_tariffs.get(selected_category, [])}
    for t in selected_tariffs:
        if t in lookup:
            added_items.append({"Category": selected_category, "Tariff": t, "Amount": lookup[t]})
    return added_items

# ---------- Streamlit UI ----------
st.title("📄 Medical Quotation Generator")

st.subheader("1. Upload Charge Sheet")
charge_file = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])

if charge_file:
    df_all = load_charge_sheet(charge_file)
    category_tariffs = extract_category_tariffs(df_all)

    st.subheader("2. Select Category")
    selected_category = st.selectbox("Category", list(category_tariffs.keys()))

    st.subheader("3. Select Tariffs (Click to Add)")
    tariffs_list = [x[0] for x in category_tariffs[selected_category]]

    selected_tariffs = st.multiselect("Tariffs", tariffs_list)

    if "quotation_items" not in st.session_state:
        st.session_state.quotation_items = []

    new_items = add_selected_tariffs(selected_category, selected_tariffs, category_tariffs)
    st.session_state.quotation_items = new_items

    st.subheader("4. Selected Tariffs")
    st.dataframe(pd.DataFrame(st.session_state.quotation_items))

    total_amt = sum([x["Amount"] for x in st.session_state.quotation_items])
    st.markdown(f"**Total Amount:** {total_amt:.2f}")

    st.subheader("5. Generate Quotation")
    st.info("Excel export logic can be added here using your template and previous fill_excel_template function.")
