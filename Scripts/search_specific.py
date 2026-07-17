import pandas as pd
import sys

file = r'folder excel\Masterlist SPS CHS 2 Layer DIG.xlsx'
try:
    xls = pd.ExcelFile(file)
    print(f"Sheets in {file}: {xls.sheet_names}")
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
        # Check if 131 exists
        mask = df.apply(lambda col: col.str.contains('131', case=False, na=False))
        if mask.any().any():
            print(f"Found '131' in {file}, Sheet: {sheet_name}")
            for col in df.columns:
                matches = df[df[col].str.contains('131', case=False, na=False) == True]
                if not matches.empty:
                    print(f"  Col '{col}': {matches[col].unique()}")
except Exception as e:
    print(f"Error: {e}")
