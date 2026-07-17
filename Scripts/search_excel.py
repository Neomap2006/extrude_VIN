import pandas as pd
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

excel_path = r'D:\Project\extrude\folder excel\Masterlist SPS CHS 2 Layer DIG.xlsx'
item_to_find = 'TA169A'

print(f"Mencari '{item_to_find}' di file: {excel_path}")

if not os.path.exists(excel_path):
    print(f"File tidak ditemukan di path: {excel_path}")
else:
    try:
        xl = pd.ExcelFile(excel_path)
        found_sheets = []
        
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            
            mask = df.astype(str).apply(lambda x: x.str.contains(item_to_find, case=False, na=False))
            if mask.any().any():
                found_sheets.append(sheet_name)
                print(f"\n[FOUND] DITEMUKAN di sheet: '{sheet_name}'")
                
                rows_with_match = mask.any(axis=1)
                matched_df = df[rows_with_match]
                
                print(f"   Baris yang cocok:")
                for idx, row in matched_df.iterrows():
                    print(f"   Baris Excel {idx+2}: {row.iloc[1:5].to_list()} ...")
                
        if not found_sheets:
            print(f"\n[NOT FOUND] Tidak ditemukan item '{item_to_find}' di sheet manapun.")
            
    except Exception as e:
        print(f"Error reading excel: {e}")
