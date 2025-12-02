if charge_sheet_file:
    try:
        xl = pd.ExcelFile(charge_sheet_file)
        st.write("Available Tabs:")
        st.write(xl.sheet_names)
    except:
        st.write("Cannot read sheet names.")
