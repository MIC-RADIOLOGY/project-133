# app.py
import streamlit as st
from utils import (
    load_tariff_sheets, find_sheet_for_scan, find_tariff_in_sheet,
    fill_template_placeholders, fill_template_by_mapping, make_output_filename, load_mapping_file
)
from pathlib import Path
import tempfile
import pandas as pd
import io
import json

st.set_page_config(page_title="Radiology Quotation Filler", layout="centered")

st.title("Radiology Quotation Auto-Filler")
st.markdown("Upload your charge sheet (Excel) and your quotation template, then enter details and generate a filled quotation.")

# ------- Upload files -------
charge_file = st.file_uploader("Upload charge sheet (Excel with tabs for USS/XRAY/CT...)", type=['xlsx','xls'], key="charge")
template_file = st.file_uploader("Upload quotation template (Excel .xlsx). Use placeholders like {{PATIENT_NAME}} OR provide mapping.json", type=['xlsx','xls'], key="template")

st.info("Two ways to fill template:\n\n1) **Placeholders**: put `{{PATIENT_NAME}}`, `{{TARIFF_DESC_1}}` etc. in your template cells.\n\n2) **Mapping**: upload a mapping.json file (see example).")

mapping_file = st.file_uploader("Optional: mapping.json (for exact cell addresses)", type=['json'], key="mapping")

if charge_file is None or template_file is None:
    st.warning("Please upload both the charge sheet and the template to continue.")
    st.stop()

# Read tariff sheets to memory (temporary file because pandas/ExcelFile needs a path-like)
with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
    tmp.write(charge_file.getbuffer())
    tmp_path = tmp.name

tariff_sheets = load_tariff_sheets(tmp_path)
st.success(f"Loaded charge sheet with sheets: {', '.join(tariff_sheets.keys())}")

# Basic inputs
patient_name = st.text_input("Patient full name", "")
medical_aid = st.text_input("Medical aid number", "")
scan_type = st.text_input("Type of scan (e.g., 'USS Abdomen', 'X-Ray Chest')", "")
code_or_lookup = st.text_input("Optional: tariff code or short description to help lookup (recommended)", "")

if st.button("Find tariff"):
    if not scan_type:
        st.error("Please enter a scan type.")
    else:
        sheet_key = find_sheet_for_scan(tariff_sheets, scan_type)
        st.write(f"Selected sheet: **{sheet_key}**")
        df = tariff_sheets[sheet_key]
        found = find_tariff_in_sheet(df, scan_type, code_or_desc=code_or_lookup)
        if found:
            st.success("Tariff found:")
            st.json(found)
        else:
            st.error("No tariff matched. Try using a better description or provide tariff code in the helper field.")
            # show first rows for inspection
            st.dataframe(df.head())

# Allow multiple tariffs
st.markdown("### Add scan items (multiple allowed)")
items = []
# Use a session state list to accumulate items
if 'items' not in st.session_state:
    st.session_state['items'] = []

col1, col2, col3 = st.columns([5,3,3])
with col1:
    item_scan = st.text_input("Scan description to add", key="item_scan")
with col2:
    item_code = st.text_input("Optional tariff code", key="item_code")
with col3:
    add_item_btn = st.button("Add item")

if add_item_btn:
    if not item_scan:
        st.error("Enter a scan description to add.")
    else:
        sheet_key = find_sheet_for_scan(tariff_sheets, item_scan)
        df = tariff_sheets.get(sheet_key)
        found = None
        if df is not None:
            found = find_tariff_in_sheet(df, item_scan, code_or_desc=item_code)
        if not found:
            st.warning("No tariff matched automatically — adding with price 0.0. You can edit price later.")
            found = {'TARIFF_CODE': item_code or '', 'DESCRIPTION': item_scan, 'PRICE': 0.0}
        st.session_state['items'].append(found)
        st.success("Item added to quote.")

if st.session_state['items']:
    st.markdown("#### Current items")
    df_items = pd.DataFrame(st.session_state['items'])
    # allow editing prices inline (streamlit doesn't support table editing natively; show and accept changes via inputs)
    st.dataframe(df_items)
    if st.button("Clear items"):
        st.session_state['items'] = []

# Choose fill method
st.markdown("### Template filling options")
fill_method = st.radio("Choose template fill method", ("Placeholders (recommended)", "Cell mapping (mapping.json)"))

# If mapping file provided, load it
mapping = None
if mapping_file is not None:
    try:
        mapping = json.loads(mapping_file.getvalue().decode('utf-8'))
        st.write("Loaded mapping.json")
    except Exception as e:
        st.error(f"Failed to parse mapping.json: {e}")
        mapping = None

if fill_method == "Cell mapping (mapping.json)" and mapping is None:
    st.info("Please upload mapping.json or switch to placeholder method.")
    st.stop()

# Generate quotation
if st.button("Generate quotation"):
    if not patient_name:
        st.error("Please enter patient name.")
        st.stop()
    if not st.session_state['items']:
        st.error("Please add at least one scan item.")
        st.stop()

    # build replacements / values
    # prepare ITEMS as TARIFF_CODE_1, TARIFF_DESC_1, TARIFF_PRICE_1 etc.
    replacements = {}
    total = 0.0
    for i, it in enumerate(st.session_state['items'], start=1):
        replacements[f"TARIFF_CODE_{i}"] = it.get('TARIFF_CODE','')
        replacements[f"TARIFF_DESC_{i}"] = it.get('DESCRIPTION','')
        replacements[f"TARIFF_PRICE_{i}"] = f"{it.get('PRICE',0.0):.2f}"
        total += float(it.get('PRICE', 0.0))
    replacements['TOTAL'] = f"{total:.2f}"
    replacements['PATIENT_NAME'] = patient_name
    replacements['MEDICAL_AID'] = medical_aid
    replacements['DATE'] = pd.Timestamp.now().strftime("%Y-%m-%d")

    # save uploaded template to temp file
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as ttmp:
        ttmp.write(template_file.getbuffer())
        tmp_template_path = ttmp.name

    out_name = make_output_filename(patient=patient_name)
    out_path = Path(tempfile.gettempdir()) / out_name

    if fill_method == "Placeholders (recommended)":
        fill_template_placeholders(tmp_template_path, out_path, replacements)
    else:
        # mapping expects dict field -> cell address
        fill_template_by_mapping(tmp_template_path, out_path, mapping, replacements)

    # present file for download
    with open(out_path, 'rb') as fh:
        data = fh.read()
    st.success("Quotation generated.")
    st.download_button("Download quotation (xlsx)", data, file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
