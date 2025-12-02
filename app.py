import pandas as pd
import glob
import os

# --- Configuration ---
# The list of CSV files created from your uploaded Excel file.
# We will use glob to find all the relevant tariff files.
# The `error_bad_lines=False` is used to handle potential inconsistencies 
# in the CSV files, especially those with many blank rows or irregular headers.
# Note: In newer pandas versions, this is replaced by `on_bad_lines='skip'`.
TARIFF_FILES = glob.glob('CIMAS - TARIFFS DECEMBER 2024.xlsx - *.csv')

# Define the columns where we expect to find the Examination description and the Tariff code.
# The actual column index might shift due to initial blank columns in the CSVs.
EXAMINATION_COL_NAMES = ['EXAMINATION', 'Description']
TARIFF_CODE_COL_NAME = 'TARIFF'
CIMAS_USD_COL_NAME = 'CIMAS USD'

# --- Data Loading and Preparation ---

def load_data():
    """Loads all relevant tariff CSVs into a single dictionary of DataFrames."""
    dataframes = {}
    print(f"Loading {len(TARIFF_FILES)} tariff sheets...")
    
    for file_path in TARIFF_FILES:
        try:
            # Sheet name is derived from the file name
            sheet_name = os.path.basename(file_path).replace('CIMAS - TARIFFS DECEMBER 2024.xlsx - ', '').replace('.csv', '')
            
            # We skip the initial few rows which seem to be header context rows
            # We use a combined approach to find the header (row 12 for most, 3 for others)
            df = pd.read_csv(file_path, encoding='utf-8', header=None, skiprows=lambda x: x < 10)
            
            # Try to dynamically detect the true header row by finding the 'EXAMINATION' or 'TARIFF' column
            header_row_index = -1
            for i in range(len(df)):
                if any(col in df.iloc[i].values for col in EXAMINATION_COL_NAMES) and TARIFF_CODE_COL_NAME in df.iloc[i].values:
                    header_row_index = i
                    break
            
            # Reload the CSV using the detected header row
            if header_row_index != -1:
                df = pd.read_csv(file_path, encoding='utf-8', header=header_row_index + 10, on_bad_lines='skip')
                
                # Standardize column names for easier querying
                cols = df.columns.tolist()
                
                # Rename the column that holds the examination name
                for old_name in EXAMINATION_COL_NAMES:
                    if old_name in cols:
                        df.rename(columns={old_name: 'Examination'}, inplace=True)
                        break
                        
                # Rename the tariff code column
                if TARIFF_CODE_COL_NAME in cols:
                    df.rename(columns={TARIFF_CODE_COL_NAME: 'Tariff Code'}, inplace=True)
                    
                # Drop rows where 'Examination' is null (usually blank rows/totals)
                if 'Examination' in df.columns:
                    df.dropna(subset=['Examination'], inplace=True)
                    # Convert 'Tariff Code' to string/int for reliable searching
                    if 'Tariff Code' in df.columns:
                         df['Tariff Code'] = pd.to_numeric(df['Tariff Code'], errors='coerce').fillna('').astype(str).str.split('.').str[0]
                         
                    dataframes[sheet_name] = df
                    print(f"Successfully loaded '{sheet_name}' with {len(df)} records.")
                else:
                    print(f"Skipping '{sheet_name}': Could not find 'EXAMINATION' column.")

            else:
                 print(f"Skipping '{sheet_name}': Could not reliably detect header row.")

        except Exception as e:
            print(f"Error loading {file_path}: {e}")
            
    return dataframes

ALL_TARIFF_DATA = load_data()


# --- Core Functions ---

def search_tariff_by_examination(keyword):
    """Searches across all dataframes for examinations matching the keyword."""
    results = []
    keyword = str(keyword).strip().lower()

    if not ALL_TARIFF_DATA:
        return "No tariff data loaded.", []

    for sheet_name, df in ALL_TARIFF_DATA.items():
        if 'Examination' in df.columns:
            # Find matching rows
            mask = df['Examination'].astype(str).str.lower().str.contains(keyword, na=False)
            matches = df[mask]
            
            if not matches.empty:
                for index, row in matches.iterrows():
                    tariff_code = row.get('Tariff Code', 'N/A')
                    examination = row['Examination']
                    cimas_usd = row.get(CIMAS_USD_COL_NAME)
                    
                    # Collect all available tariff columns (excluding Examination, Tariff Code, QTY etc.)
                    tariffs = {
                        col: row[col] 
                        for col in df.columns 
                        if col not in ['Examination', 'Tariff Code', 'QTY', 'MOD', 'HRS'] and pd.notna(row[col])
                    }

                    results.append({
                        'Sheet': sheet_name,
                        'Examination': examination,
                        'Tariff Code': tariff_code,
                        'CIMAS USD Price': f"${cimas_usd:,.2f}" if pd.notna(cimas_usd) and pd.to_numeric(cimas_usd, errors='coerce') is not None else 'N/A',
                        'All Tariffs': tariffs
                    })
    
    return f"Found {len(results)} results for '{keyword}'.", results


def search_examination_by_tariff_code(code):
    """Searches across all dataframes for a specific tariff code."""
    results = []
    code_str = str(code).strip()

    if not ALL_TARIFF_DATA:
        return "No tariff data loaded.", []

    for sheet_name, df in ALL_TARIFF_DATA.items():
        if 'Tariff Code' in df.columns:
            # Ensure the Tariff Code column is clean strings
            mask = df['Tariff Code'].astype(str).str.strip() == code_str
            matches = df[mask]
            
            if not matches.empty:
                for index, row in matches.iterrows():
                    examination = row.get('Examination', 'N/A')
                    cimas_usd = row.get(CIMAS_USD_COL_NAME)

                    tariffs = {
                        col: row[col] 
                        for col in df.columns 
                        if col not in ['Examination', 'Tariff Code', 'QTY', 'MOD', 'HRS'] and pd.notna(row[col])
                    }

                    results.append({
                        'Sheet': sheet_name,
                        'Examination': examination,
                        'Tariff Code': code_str,
                        'CIMAS USD Price': f"${cimas_usd:,.2f}" if pd.notna(cimas_usd) and pd.to_numeric(cimas_usd, errors='coerce') is not None else 'N/A',
                        'All Tariffs': tariffs
                    })

    return f"Found {len(results)} examinations for tariff code '{code_str}'.", results


def summarize_results(results):
    """Formats the raw results into a readable string."""
    if not results:
        return "No details found."
        
    output = []
    for item in results:
        output.append(f"\n--- Examination: {item['Examination']} (Tariff Code: {item['Tariff Code']}) ---")
        output.append(f"  Source Sheet: {item['Sheet']}")
        output.append(f"  CIMAS USD Rate: {item['CIMAS USD Price']}")
        
        # Displaying a subset of the available plan tariffs
        tariff_display = [f"{k}: {v:,.2f}" if pd.to_numeric(v, errors='coerce') is not None else f"{k}: {v}" 
                          for k, v in item['All Tariffs'].items() if k != CIMAS_USD_COL_NAME]
        
        # Only show the first 5 tariffs for brevity in the summary
        output.append("  Sample Tariffs:")
        output.extend([f"    - {t}" for t in tariff_display[:5]])
        if len(tariff_display) > 5:
            output.append(f"    - ... and {len(tariff_display) - 5} more plans.")
            
    return "\n".join(output)


# --- Example Usage ---

if __name__ == "__main__":
    print("\n--- Running Tariff Analysis Examples ---\n")

    # Example 1: Search by Examination Keyword
    keyword = "CT Head"
    message, results = search_tariff_by_examination(keyword)
    print(f"Query: Search for '{keyword}'")
    print(message)
    print(summarize_results(results))
    print("\n" + "="*50 + "\n")

    # Example 2: Search by a specific Tariff Code
    tariff_code = "76925"
    message, results = search_examination_by_tariff_code(tariff_code)
    print(f"Query: Search for Tariff Code '{tariff_code}'")
    print(message)
    print(summarize_results(results))
    print("\n" + "="*50 + "\n")
    
    # Example 3: Search for a different examination
    keyword = "Mammography"
    message, results = search_tariff_by_examination(keyword)
    print(f"Query: Search for '{keyword}'")
    print(message)
    print(summarize_results(results))
    print("\n" + "="*50 + "\n")
