import streamlit as st
import pandas as pd
import openpyxl
import io
from copy import copy
from typing import Optional

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ---------- Config / heuristics ----------
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
                "MODIFIER": clean_text(r["C_MOD"]),
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

# ---------- Streamlit UI ----------
st.title("Medical Quotation Generator — Category / Tariff Selection (MRI supported)")

debug_mode = st.checkbox("Show parsing debug output", value=False)

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if uploaded_charge:
    if st.button("Load & Parse Charge Sheet"):
        parsed = load_charge_sheet(uploaded_charge)
        st.session_state.parsed_df = parsed
        st.session_state.selected_rows = []
        st.success("Charge sheet parsed successfully.")

if "parsed_df" in st.session_state:
    df = st.session_state.parsed_df

    if debug_mode:
        st.dataframe(df.head(200))

    cats = [c for c in sorted(df["CATEGORY"].dropna().unique())] if "CATEGORY" in df.columns else []

    if cats:
        main_sel = st.selectbox("Select Main Category", ["-- choose --"] + cats)
        if main_sel != "-- choose --":
            subs = [s for s in sorted(df[df["CATEGORY"] == main_sel]["SUBCATEGORY"].dropna().unique())]
            sub_sel = st.selectbox("Select Subcategory", ["-- all --"] + subs) if subs else "-- all --"

            if sub_sel == "-- all --":
                scans_for_cat = df[df["CATEGORY"] == main_sel].reset_index(drop=True)
            else:
                scans_for_cat = df[(df["CATEGORY"] == main_sel) & (df["SUBCATEGORY"] == sub_sel)].reset_index(drop=True)

            if not scans_for_cat.empty:
                scans_for_cat["label"] = scans_for_cat.apply(
                    lambda r: f"{r['SCAN']} | Tariff: {r['TARIFF']} | Mod: {r['MODIFIER']} | Amt: {r['AMOUNT']}",
                    axis=1
                )

                selected_indices = st.multiselect("Select Tariffs", options=list(range(len(scans_for_cat))),
                                                format_func=lambda i: scans_for_cat.at[i, "label"])

                # Auto-add selected tariffs without add button
                st.session_state.selected_rows = [scans_for_cat.iloc[i].to_dict() for i in selected_indices]

    st.subheader("Selected Items")
    if "selected_rows" in st.session_state and st.session_state.selected_rows:
        sel_df = pd.DataFrame(st.session_state.selected_rows)
        st.dataframe(sel_df[["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]].reset_index(drop=True))
        total_amt = sum([safe_float(r.get("AMOUNT", 0.0), 0.0) for r in st.session_state.selected_rows])
        st.markdown(f"**Total Amount:** {total_amt:.2f}")
    else:
        st.info("Select tariffs from the list above.")
