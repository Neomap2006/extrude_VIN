import pandas as pd
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

folder_path = r'D:\Project\extrude\folder excel'
item_to_find = 'TA169A'

print(f"Mencari '{item_to_find}' di semua file excel dalam folder: {folder_path}")

if not os.path.exists(folder_path):
    print(f"Folder tidak ditemukan!")
else:
    found_anywhere = False
    
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') and not filename.startswith('~'):
            excel_path = os.path.join(folder_path, filename)
            try:
                xl = pd.ExcelFile(excel_path)
                for sheet_name in xl.sheet_names:
                    df = xl.parse(sheet_name)
                    mask = df.astype(str).apply(lambda x: x.str.contains(item_to_find, case=False, na=False))
                    
                    if mask.any().any():
                        found_anywhere = True
                        print(f"\n[FOUND] DITEMUKAN di file: '{filename}', sheet: '{sheet_name}'")
                        
                        rows_with_match = mask.any(axis=1)
                        matched_df = df[rows_with_match]
                        
                        for idx, row in matched_df.iterrows():
                            # Show doc number column which is usually index 3 or 4
                            print(f"   Baris Excel {idx+2}: {row.iloc[:5].to_list()} ...")
            except Exception as e:
                print(f"Error reading {filename}: {e}")
                
    if not found_anywhere:
        print(f"\n[NOT FOUND] Item '{item_to_find}' tidak ditemukan di SELURUH file Excel di folder tersebut.")
