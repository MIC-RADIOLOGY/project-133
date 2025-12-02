import streamlit as st
import pandas as pd
import io
import re

# List of all relevant CIMAS tariff files to load from the environment
# These file names are derived from the uploaded CIMAS spreadsheet sections.
CIMAS_FILES = [
    "CIMAS - TARIFFS DECEMBER 2024.xlsx - IMAGE INTESIFIER .csv",
    "CIMAS - TARIFFS DECEMBER 2024.xlsx - MAMMOGRAPHY.csv",
    "CIMAS - TARIFFS DECEMBER 2024.xlsx - FLUROSCOPY RECIPES.csv",
    "CIMAS - TARIFFS DECEMBER 2024.xlsx - CT ANGIO RECIPES.csv",
    "CIMAS - TARIFFS DECEMBER 2024.xlsx - CT SCAN RECIPES.csv",
    "CIMAS - TARIFFS DECEMBER 2024.xlsx - XRAY RECIPES.csv",
    "CIMAS - TARIFFS DECEMBER 2024.xlsx - USS RECIPES.csv",
    "CIMAS - TARIFFS DECEMBER 2024.xlsx - USS DOPPLERS.csv",
]

# Function to attempt to load and clean the CSV data
@st.cache_data
def load_and_clean_data(file_paths):
    """
    Loads multiple CSV files, attempts to auto-detect the header row, 
    standardizes column names, and combines the data into a single DataFrame.
    
    This function is cached by Streamlit to avoid reprocessing on every rerun.
    """
    all_data = []
    
    # Define potential names for the key columns (case-insensitive matching)
    EXAMINATION_KEYS = ['EXAMINATION', 'Description']
    TARIFF_KEYS = ['TARIFF', 'Tariff ']
    # Look for common USD-related column names
    USD_KEYS = ['CIMAS USD', 'USD', 'UNIT USD', 'FEES', 'AMOUNT'] 
    
    # Pre-compiled regex for filtering out noise (like "FF", "Total", blank lines)
    # This helps clean up rows that are often subtotals or filler data
    noise_pattern = re.compile(r'^(ff|total|co-?payment|exam\s*total|\s*)$', re.IGNORECASE)

    for file_path in file_paths:
        try:
            # 1. Read the file, skipping no rows initially, to detect the header
            df_raw = pd.read_csv(file_path, header=None, skiprows=0, keep_default_na=True)
            
            # 2. Find the row containing the key header words (EXAMINATION or TARIFF)
            header_row_index = -1
            for i, row in df_raw.iterrows():
                # Check for any key column name match in the row values
                row_str = ' '.join(str(x) for x in row.values if pd.notna(x)).upper()
                if any(key.strip().upper() in row_str for key in EXAMINATION_KEYS + TARIFF_KEYS):
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                st.warning(f"Skipping {file_path}: Could not reliably determine header row.")
                continue

            # 3. Reload the dataframe using the correct header row
            df = pd.read_csv(file_path, header=header_row_index)
            
            # 4. Normalize column names
            col_map = {}
            
            # List of standard columns to look for
            standard_cols = {
                'Examination': EXAMINATION_KEYS,
                'Tariff Code': TARIFF_KEYS,
                'CIMAS USD': USD_KEYS
            }

            # Map found original columns to standard names
            for standard_name, potential_names in standard_cols.items():
                found_col = next(
                    (col for col in df.columns if str(col).strip().upper() in (name.strip().upper() for name in potential_names)), 
                    None
                )
                if found_col:
                    col_map[found_col] = standard_name
                # If a required column is not found, it will be added to the missing list below.

            required_standard_names = list(standard_cols.keys())
            
            # Check if all required columns were mapped (this is crucial for data integrity)
            if len(col_map) < len(required_standard_names):
                missing = [name for name in required_standard_names if name not in col_map.values()]
                st.warning(f"Skipping {file_path}: Missing key columns in header row: {missing}. Found: {list(col_map.values())}")
                continue


            # 5. Select, rename, and clean data
            df_clean = df[list(col_map.keys())].rename(columns=col_map)
            
            # Clean Examination column: remove NaNs and filter noise
            df_clean['Examination'] = df_clean['Examination'].astype(str).str.strip()
            df_clean = df_clean[
                df_clean['Examination'].notna() & 
                (df_clean['Examination'].str.len() > 3) & # Filter out very short, often non-descriptive strings
                (~df_clean['Examination'].apply(lambda x: bool(noise_pattern.match(x))))
            ].copy()
            
            # Ensure Tariff Code is clean
            df_clean['Tariff Code'] = df_clean['Tariff Code'].astype(str).str.strip()

            # Convert CIMAS USD to numeric, coercing errors to NaN
            df_clean['CIMAS USD'] = pd.to_numeric(df_clean['CIMAS USD'], errors='coerce')
            
            # Remove rows where the USD value is invalid, NaN, or 0
            df_clean = df_clean[df_clean['CIMAS USD'].notna() & (df_clean['CIMAS USD'] > 0)]

            # Add source file name for traceability
            source_name = file_path.split(' - ')[-1].replace('.csv', '').strip()
            df_clean['Source'] = source_name
            
            all_data.append(df_clean)

        except Exception as e:
            st.error(f"Error processing {file_path}: {e}")

    if all_data:
        # Combine all processed dataframes
        combined_df = pd.concat(all_data, ignore_index=True)
        # Drop duplicates based on the primary fields (Examination and Code)
        combined_df = combined_df.drop_duplicates(subset=['Examination', 'Tariff Code'], keep='first')
        combined_df = combined_df.reset_index(drop=True)
        return combined_df
    else:
        return pd.DataFrame()

# --- Main Application Function ---

def main():
    # Set the page configuration for a wider layout
    st.set_page_config(layout="wide", page_title="Medical Tariff Search")

    # Load data
    df_tariffs = load_and_clean_data(CIMAS_FILES)
    
    # --- Application Title and Description ---
    st.title("🏥 Medical Imaging Tariff Search")
    st.markdown("Use this tool to search for medical tariffs (Examination Name or TARIFF Code) across the combined CIMAS data sheets.")
    
    if df_tariffs.empty:
        st.error("❌ Failed to load and combine tariff data from the CSV files. Please check the file formats or console for warnings.")
        return

    st.success(f"✅ Successfully loaded **{len(df_tariffs)}** unique tariff items from **{len(CIMAS_FILES)}** source files.")

    # --- Search Input ---
    search_query = st.text_input(
        "Enter Examination Name or TARIFF Code:",
        placeholder="e.g., Head, CT, Pelvis, 77001",
        key="search_input"
    ).strip()

    # --- Search Logic ---
    if search_query:
        # Convert the query to lowercase for case-insensitive search
        query_lower = search_query.lower()
        
        # 1. Search by Examination Name (case-insensitive, contains)
        examination_match = df_tariffs['Examination'].astype(str).str.lower().str.contains(query_lower, na=False)
        
        # 2. Search by Tariff Code (contains for string/number search)
        tariff_match = df_tariffs['Tariff Code'].astype(str).str.contains(search_query, na=False)
        
        # Combine the matches
        filtered_df = df_tariffs[examination_match | tariff_match].copy()

        if not filtered_df.empty:
            st.subheader(f"Search Results for '{search_query}' ({len(filtered_df)} items)")
            
            # Format the USD column for better display
            filtered_df['CIMAS USD'] = filtered_df['CIMAS USD'].apply(lambda x: f"${x:,.2f}")
            
            # Select and reorder columns for display
            display_df = filtered_df[['Examination', 'Tariff Code', 'CIMAS USD', 'Source']]
            
            # Display the results in an interactive table
            st.dataframe(
                display_df, 
                use_container_width=True,
                # Define column configurations for width and titles
                column_config={
                    "Examination": st.column_config.TextColumn("Examination", width="large"),
                    "Tariff Code": st.column_config.TextColumn("Tariff Code", width="small"),
                    "CIMAS USD": st.column_config.TextColumn("USD Tariff", width="small"),
                    "Source": st.column_config.TextColumn("Source Sheet", width="medium"),
                },
                hide_index=True
            )
        else:
            st.warning(f"No results found for **'{search_query}'**. Try a different or broader term.")
            
    else:
        st.info("Start typing an examination name (e.g., 'USS', 'CT Head') or a tariff code (e.g., '77002') to search.")
        
        # Optional: Show a summary of loaded data when no search is active
        st.subheader("Data Source Summary")
        source_counts = df_tariffs['Source'].value_counts().reset_index()
        source_counts.columns = ['Source Sheet', 'Number of Tariffs']
        st.markdown("Tariff counts per original spreadsheet sheet:")
        st.dataframe(source_counts, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
