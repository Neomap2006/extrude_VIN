import pandas as pd

import glob
import warnings
warnings.simplefilter("ignore")

excel_files = glob.glob(r"d:\Project\extrude\folder excel\Masterlist*.xlsx")
doc = 'SOP/PROD/HOSE/SPS/2/9'

for file_path in excel_files:
    if "~$" in file_path: continue
    print(f"\n--- Searching in {file_path.split('\\\\')[-1]} ---")
    try:
        df = pd.read_excel(file_path, sheet_name=0, header=2)
        df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
        
        mask = df.apply(lambda row: row.astype(str).str.contains(doc, case=False).any(), axis=1)
        matching_rows = df[mask]
        
        if not matching_rows.empty:
            print(f"✅ FOUND in {file_path.split('\\\\')[-1]}!")
            doc_col = [c for c in df.columns if 'NO' in c.upper() or 'DOC' in c.upper() or 'SPS' in c.upper()]
            spiral_cols = [c for c in df.columns if 'SPIRAL' in str(c).upper() or 'SPEED' in str(c).upper()]
            
            for _, row in matching_rows.iterrows():
                print("\n  DATA:")
                for col in ['ITEM CODE', 'DIMENSI'] + doc_col[:3] + spiral_cols:
                    if col in df.columns:
                        print(f"  {col:25} : {row[col]}")
    except Exception as e:
        print(f"  Error reading file: {e}")
