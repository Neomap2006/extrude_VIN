import pandas as pd
import glob
import sys

files = glob.glob(r'folder excel\*.xlsx')
for file in files:
    try:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
            # check if 131 exists anywhere
            mask = df.apply(lambda col: col.str.contains('131', case=False, na=False))
            if mask.any().any():
                rows, cols = mask.where(mask).dropna(how='all').dropna(axis=1, how='all').shape
                print(f"Found '131' in {file}, Sheet: {sheet_name} ({rows} rows matched)")
                # Print exact document names matching
                for col in df.columns:
                    matches = df[df[col].str.contains('131', case=False, na=False) == True]
                    if not matches.empty:
                        print(f"  Col '{col}': {matches[col].unique()}")
    except Exception as e:
        print(f"Error reading {file}: {e}")
