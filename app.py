
def main():
    """Main function to run the Streamlit application."""
    st.set_page_config(layout="wide", page_title="Medical Tariff Search")

    st.title("🏥 Medical Tariff Data Analyzer")
    st.markdown("A tool to search and compare examination tariffs across different schemes from the provided CSV files.")
    
    # --- Sidebar for Data Info ---
    st.sidebar.header("Data Status")
    if ALL_TARIFF_DATA:
        st.sidebar.markdown(f"**{len(ALL_TARIFF_DATA)}** Sheets Loaded:")
    else:
        st.sidebar.error("No data could be loaded. Please ensure CSV files are in the same directory.")
        return # Stop execution if no data is available
        
    # --- Search Input ---
    search_by = st.radio(
        "Search By:", 
        ('Examination Name', 'Tariff Code'), 
        key='search_mode', 
        horizontal=True
    )
    
    if search_by == 'Examination Name':
        search_query = st.text_input(
            "Enter part of the Examination Name (e.g., 'CT Head', 'USS')", 
            key='examination_query'
        )
    else:
        search_query = st.text_input(
            "Enter the exact 5-digit Tariff Code (e.g., '77001', '76925')", 
            key='tariff_code_query'
        ).strip()
    
    
    # --- Perform Search and Display Results ---
    if search_query:
        with st.spinner(f"Searching for '{search_query}'..."):
            mode = 'examination' if search_by == 'Examination Name' else 'tariff_code'
            results = search_tariff_data(search_query, mode)
        
        st.subheader(f"Found {len(results)} Matches")

        if results:
            # 1. Prepare a summary table
            summary_data = []
            for item in results:
                summary_data.append({
                    'Examination': item['Examination'],
                    'Tariff Code': item['Tariff Code'],
                    'Sheet': item['Sheet'],
                    'CIMAS USD': item['CIMAS USD Price']
                })
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

            # 2. Detailed view using expanders
            st.markdown("---")
            st.subheader("Detailed Breakdown")
            for i, item in enumerate(results):
                header = f"{item['Examination']} (Code: {item['Tariff Code']}) — Sheet: {item['Sheet']}"
                with st.expander(header):
                    st.metric("CIMAS USD Rate", item['CIMAS USD Price'])
                    
                    # Convert tariffs dictionary to a DataFrame for clean display
                    tariff_items = [
                        {'Scheme/Plan': k, 'Rate': v} 
                        for k, v in item['All Tariffs'].items()
                    ]
                    
                    # Convert Rates to numeric if possible for better sorting/display
                    tariff_df = pd.DataFrame(tariff_items)
                    tariff_df['Rate'] = pd.to_numeric(tariff_df['Rate'], errors='coerce')
                    
                    # Display the full list of rates from the source sheet
                    st.dataframe(
                        tariff_df.sort_values(by='Rate', ascending=False, na_position='last'), 
                        use_container_width=True, 
                        hide_index=True
                    )
        else:
            st.info("No matching examinations or tariff codes found.")

if __name__ == "__main__":
    main()
