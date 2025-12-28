# app.py
import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime

st.set_page_config(page_title="Medical Quotation Generator", layout="wide")

# ------------------------------------------------------------
# LOGIN / CAPTIVE PORTAL (SAFE)
# ------------------------------------------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("Login Required")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    
    login_attempted = st.button("Login")

    if login_attempted:
        # Replace with your credentials
        if username == "admin" and password == "Jamela2003":
            st.session_state.logged_in = True
            st.success("Login successful! Reload or interact with the app to continue.")
        else:
            st.error("Invalid credentials")
    
    st.stop()  # stop execution until login succeeds

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
COMPONENT_KEYS = {
    "PELVIS", "CONSUMABLES", "FF",
    "IV", "IV CONTRAST", "IV CONTRAST 100MLS"
}

GARBAGE_KEYS = {"TOTAL", "CO-PAYMENT", "CO PAYMENT", "CO - PAYMENT", "CO", ""}

MAIN_CATEGORIES = set()

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\xa0", " ").strip()

def safe_int(x, default=1):
    try:
        return int(float(str(x).replace(",", "").strip()))
    except Exception:
        return default

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", "").strip())
    except Exception:
        return default

# ------------------------------------------------------------
# PARSER
# ------------------------------------------------------------
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

        exam_u = exam.upper().strip()

        if exam_u in MAIN_CATEGORIES or exam_u.endswith("SCAN") or exam_u in {"XRAY", "MRI", "ULTRASOUND"}:
            MAIN_CATEGORIES.add(exam_u)
            current_category = exam
            current_subcategory = None
            continue

        if exam_u in GARBAGE_KEYS:
            continue

        if clean_text(r["B_TARIFF"]) == "" and clean_text(r["E_AMOUNT"]) == "":
            current_subcategory = exam
            continue

        if not current_category:
            continue

        is_main_scan = exam_u not in COMPONENT_KEYS

        structured.append({
            "CATEGORY": current_category,
            "SUBCATEGORY": current_subcategory,
            "SCAN": exam,
            "IS_MAIN_SCAN": is_main_scan,
            "TARIFF": safe_float(r["B_TARIFF"], None),
            "MODIFIER": str(clean_text(r["C_MOD"])),
            "QTY": safe_int(r["D_QTY"], 1),
            "AMOUNT": safe_float(r["E_AMOUNT"], 0.0)
        })

    return pd.DataFrame(structured)

# ------------------------------------------------------------
# LOAD GOOGLE SHEETS AUTOMATICALLY
# ------------------------------------------------------------
@st.cache_data
def fetch_charge_sheet():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTE25A4jR0lf4DZJpo3u6OymUHk7wF1FPMO73d0AIUvoGzdIojubtmT-dkbp8WUlw/pub?output=xlsx"
    return load_charge_sheet(url)

@st.cache_data
def fetch_quote_sheet():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRzzNViIswGXCQ8MZQyCWpx-X6h4rnTFXK87viUkfSr1XXUcC4CoVg6OPBnYV-0bQ/pub?output=xlsx"
    df = pd.read_excel(url)
    return df

# ------------------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------------------
st.title("Medical Quotation Generator")

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

# Load charge sheet automatically
if st.button("Load Charge Sheet from Google Sheets"):
    st.session_state.df = fetch_charge_sheet()
    st.success("Charge sheet loaded successfully from Google Sheets!")

# If charge sheet is loaded
if "df" in st.session_state:
    df = st.session_state.df

    main_sel = st.selectbox(
        "Select Main Category",
        sorted(df["CATEGORY"].dropna().unique())
    )

    subcats = sorted(df[df["CATEGORY"] == main_sel]["SUBCATEGORY"].dropna().unique())
    sub_sel = st.selectbox("Select Subcategory", subcats) if subcats else None

    scans = (
        df[(df["CATEGORY"] == main_sel) & (df["SUBCATEGORY"] == sub_sel)]
        if sub_sel else df[df["CATEGORY"] == main_sel]
    ).reset_index(drop=True)

    scans["label"] = scans.apply(
        lambda r: f"{r['SCAN']} | Tariff {r['TARIFF']} | Amount {r['AMOUNT']}",
        axis=1
    )

    selected = st.multiselect(
        "Select scans to include",
        options=list(range(len(scans))),
        format_func=lambda i: scans.at[i, "label"]
    )

    selected_rows = [scans.iloc[i].to_dict() for i in selected]

    if selected_rows:
        st.subheader("Edit final description for Excel")
        for i, row in enumerate(selected_rows):
            new_desc = st.text_input(
                f"Description for '{row['SCAN']}'",
                value=row['SCAN'],
                key=f"desc_{i}"
            )
            selected_rows[i]['SCAN'] = new_desc

        st.subheader("Preview of selected scans")
        st.dataframe(pd.DataFrame(selected_rows)[
            ["SCAN", "TARIFF", "MODIFIER", "QTY", "AMOUNT"]
        ])

        # Load quote template from Google Sheets automatically
        if st.button("Generate & Download Quotation"):
            quote_df = fetch_quote_sheet()  # Fetch the quote sheet if needed

            # Here you can integrate quote_df as needed, for example, to enrich scan info
            # Currently it uses the same selected_rows for template filling

            safe_name = "".join(
                c for c in (patient or "patient")
                if c.isalnum() or c in (" ", "_")
            ).strip()

            # Use last uploaded template OR your quote_df as template if needed
            st.warning("Using uploaded template is recommended for correct formatting!")
            uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

            if uploaded_template:
                from app_helpers import fill_excel_template  # reuse existing function
                out = fill_excel_template(
                    uploaded_template, patient, member, provider, selected_rows
                )

                st.download_button(
                    "Download Quotation",
                    data=out,
                    file_name=f"quotation_{safe_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
