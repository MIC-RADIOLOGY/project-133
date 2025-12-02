import pandas as pd
import streamlit as st

st.set_page_config(page_title="Medical Examination Tariff Calculator", layout="wide")
st.title("Medical Examination Tariff Calculator 📊")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx", "xls"])
if uploaded_file is not None:
    # Load Excel file with all sheets
    xls = pd.ExcelFile(uploaded_file)
    st.sidebar.subheader("Select Sheet")
    sheet_name = st.sidebar.selectbox("Choose a sheet to view", xls.sheet_names)
    
    # Load the selected sheet
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    st.subheader(f"Raw Data from sheet: {sheet_name}")
    st.dataframe(df)

    # Clean column names
    df.columns = df.columns.str.strip()

    # Convert numeric columns
    numeric_cols = ['TARIFF', 'QTY', 'STANDARD', 'COMPREHENSIVE', 'CIMAS USD']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Calculate totals per EXAMINATION
    if 'EXAMINATION' in df.columns:
        st.subheader("Totals per Examination")
        totals = df.groupby('EXAMINATION')[['STANDARD', 'COMPREHENSIVE', 'CIMAS USD']].sum()
        st.dataframe(totals)

        st.subheader("Grand Totals")
        grand_totals = totals.sum()
        st.write(grand_totals)

        # Download button
        totals_csv = totals.to_csv().encode('utf-8')
        st.download_button(
            label="Download Totals as CSV",
            data=totals_csv,
            file_name=f"{sheet_name}_totals.csv",
            mime="text/csv"
        )

    st.info("✅ Upload another sheet to continue analyzing.")

st.markdown("---")
st.markdown("Developed with ❤️ using Streamlit and Pandas")
