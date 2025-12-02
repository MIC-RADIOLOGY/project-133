import pandas as pd
import sys

class MedicalTariff:
    def __init__(self, excel_path):
        # Load all sheets from the Excel file
        self.data = pd.read_excel(excel_path, sheet_name=None)
        # Clean each sheet
        for sheet, df in self.data.items():
            self.data[sheet] = self._clean_df(df)

    def _clean_df(self, df):
        # Remove empty rows and columns
        df = df.dropna(how='all').reset_index(drop=True)
        df = df.loc[:, df.notna().any()]
        # Clean column names by stripping spaces
        df.columns = [str(col).strip() for col in df.columns]
        # Try converting STANDARD, COMPREHENSIVE, CIMAS USD to numeric (if present)
        for col in ['STANDARD', 'COMPREHENSIVE', 'CIMAS USD']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        return df

    def list_sheets(self):
        return list(self.data.keys())

    def get_sheet_data(self, sheet_name):
        return self.data.get(sheet_name)

    def get_total_per_exam(self, sheet_name):
        df = self.get_sheet_data(sheet_name)
        if df is None:
            raise ValueError(f"Sheet '{sheet_name}' not found.")
        required_cols = ['EXAMINATION', 'STANDARD', 'COMPREHENSIVE', 'CIMAS USD']
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            raise ValueError(f"Sheet '{sheet_name}' missing required columns: {missing}")
        # Group by EXAMINATION and sum numeric columns
        totals = df.groupby('EXAMINATION')[['STANDARD', 'COMPREHENSIVE', 'CIMAS USD']].sum(min_count=1)
        return totals.reset_index()

    def find_tariff_by_code(self, code):
        results = []
        for sheet, df in self.data.items():
            if 'TARIFF' in df.columns:
                matches = df[df['TARIFF'].astype(str) == str(code)]
                if not matches.empty:
                    results.append((sheet, matches))
        return results

def print_help():
    print("""
Usage: python
