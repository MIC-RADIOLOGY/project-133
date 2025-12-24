import streamlit as st
import pandas as pd
from datetime import datetime
import io
import xlwings as xw

# ---------------------------
# Keep your existing parser / helpers
# ---------------------------
# (Assume load_charge_sheet and other helper functions are unchanged)

# ---------------------------
# XLWINGS TEMPLATE FILL
# ---------------------------
def fill_excel_template_xlwings(template_file, patient, member, provider, scan_rows):
    app = xw.App(visible=False)
    wb = xw.Book(template_file)
    ws = wb.sheets[0]

    def find_cell(containing):
        for cell in ws.used_range:
            if cell.value and containing.upper() in str(cell.value).upper():
                return cell
        return None

    patient_cell = find_cell("PATIENT")
    member_cell = find_cell("MEMBER")
    provider_cell = find_cell("PROVIDER")
    date_cell = find_cell("DATE")

    if patient_cell:
        patient_cell.offset(1, 0).value = patient
    if member_cell:
        member_cell.offset(1, 0).value = member
    if provider_cell:
        provider_cell.offset(1, 0).value = provider
    if date_cell:
        date_cell.offset(1, 0).value = datetime.today().strftime("%d/%m/%Y")

    headers = ["DESCRIPTION", "TARIFF", "TARRIF", "MOD", "QTY", "FEES", "AMOUNT"]
    table_start_row = 22
    for cell in ws.used_range:
        if cell.value and any(h in str(cell.value).upper() for h in headers):
            table_start_row = cell.row + 1
            break

    col_map = {}
    for cell in ws.used_range:
        if cell.value:
            t = str(cell.value).upper()
            for h in headers:
                if h in t:
                    col_map[h] = cell.column

    grand_total = 0.0
    rowptr = table_start_row
    for sr in scan_rows:
        desc = sr["SCAN"] if sr["IS_MAIN_SCAN"] else "   " + sr["SCAN"]
        ws.cells(rowptr, col_map.get("DESCRIPTION")).value = desc
        tariff_col = col_map.get("TARIFF") or col_map.get("TARRIF")
        ws.cells(rowptr, tariff_col).value = sr["TARIFF"]
        ws.cells(rowptr, col_map.get("MOD")).value = sr["MODIFIER"]
        ws.cells(rowptr, col_map.get("QTY")).value = sr["QTY"]

        fees = sr["AMOUNT"] / sr["QTY"] if sr["QTY"] else sr["AMOUNT"]
        ws.cells(rowptr, col_map.get("FEES")).value = round(fees, 2)

        grand_total += sr["AMOUNT"]
        rowptr += 1

    if "AMOUNT" in col_map:
        ws.cells(table_start_row, col_map["AMOUNT"]).value = round(grand_total, 2)

    # Save to BytesIO
    temp_file = f"temp_{datetime.now().timestamp()}.xlsx"
    wb.save(temp_file)
    wb.close()
    app.quit()

    with open(temp_file, "rb") as f:
        out = io.BytesIO(f.read())
    out.seek(0)
    return out

# ---------------------------
# STREAMLIT UI
# ---------------------------
st.title("Medical Quotation Generator")

uploaded_charge = st.file_uploader("Upload Charge Sheet (Excel)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Quotation Template (Excel)", type=["xlsx"])

patient = st.text_input("Patient Name")
member = st.text_input("Medical Aid / Member Number")
provider = st.text_input("Medical Aid Provider", value="CIMAS")

if uploaded_charge and st.button("Load & Parse Charge Sheet"):
    st.session_state.df = load_charge_sheet(uploaded_charge)
    st.success("Charge sheet parsed successfully.")

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

        if uploaded_template and st.button("Generate & Download Quotation"):
            safe_name = "".join(
                c for c in (patient or "patient")
                if c.isalnum() or c in (" ", "_")
            ).strip()

            out = fill_excel_template_xlwings(
                uploaded_template, patient, member, provider, selected_rows
            )

            st.download_button(
                "Download Quotation",
                data=out,
                file_name=f"quotation_{safe_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
