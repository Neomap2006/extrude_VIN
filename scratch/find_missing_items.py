import pandas as pd
import glob
import os
import re
import warnings

warnings.simplefilter("ignore")

missing_parts = ['HN0490', 'HN0590', 'HN0540', 'HN0530', 'PM1110', 'HN0300', 'HN0550', 'HN0250', 'HN0560']
missing_parts = [p.upper() for p in missing_parts]

folder_path = r"d:\Project\extrude\folder excel"
files_to_check = glob.glob(os.path.join(folder_path, "Masterlist*.xlsx"))

print(f"Mencari 9 item di {len(files_to_check)} file Excel Masterlist...")

found_items = {}

def get_document_number(df):
    # Di masterlist lama, biasanya doc number ada di B1, C1, B2, atau C2
    # Di Masterlist_Digitalisasi_Cpmbined_Fix, mungkin ada di header atau kolom tertentu
    for i in range(5):
        if i >= len(df): break
        row_vals = [str(x) for x in df.iloc[i].tolist()]
        row_str = " ".join(row_vals).upper()
        if 'SOP/PROD' in row_str:
            # Extract doc number
            matches = re.findall(r'SOP/PROD/HOSE/[\w/]+', row_str)
            if matches: return matches[0]
            
            # Check cell by cell if regex didn't catch perfectly due to spacing
            for val in df.iloc[i].astype(str):
                if 'SOP/PROD' in val.upper():
                    return val.strip()
    return None

for f in files_to_check:
    if '~$' in f: continue
    
    file_name = os.path.basename(f)
    try:
        # Load all sheets
        xl = pd.ExcelFile(f)
        
        # If it's the combined fix, it might be a flat table
        if 'combined' in file_name.lower() or 'cpmbined' in file_name.lower():
            df = pd.read_excel(f, sheet_name=0)
            
            # Cari kolom Part Name/Item List
            part_col = None
            doc_col = None
            for c in df.columns:
                c_str = str(c).upper()
                if 'PART' in c_str or 'ITEM' in c_str: part_col = c
                if 'DOC' in c_str or 'SOP' in c_str: doc_col = c
                
            if part_col and doc_col:
                for idx, row in df.iterrows():
                    val = str(row[part_col]).strip().upper()
                    if val in missing_parts:
                        doc = str(row[doc_col]).strip()
                        if doc and 'SOP/PROD' in doc.upper():
                            found_items[val] = {'doc': doc, 'file': file_name, 'sheet': xl.sheet_names[0]}
                            
        # For standard masterlists (one doc per sheet or multiple)
        for sheet in xl.sheet_names:
            df = pd.read_excel(f, sheet_name=sheet, header=None)
            doc_num = get_document_number(df)
            
            # Scan entire sheet for the missing parts
            for col in df.columns:
                vals = df[col].astype(str).str.strip().str.upper().tolist()
                for part in missing_parts:
                    if part in vals and part not in found_items:
                        if doc_num:
                            found_items[part] = {'doc': doc_num, 'file': file_name, 'sheet': sheet}
                        else:
                            # Try to look around for document number
                            found_items[part] = {'doc': 'UNKNOWN (Manual Check Needed)', 'file': file_name, 'sheet': sheet}
    except Exception as e:
        print(f"Error reading {file_name}: {e}")

print("\n=== HASIL PENCARIAN ===")
for part in missing_parts:
    if part in found_items:
        data = found_items[part]
        print(f"[{part}] DITEMUKAN di {data['file']} (Sheet: {data['sheet']}) -> Document: {data['doc']}")
    else:
        print(f"[{part}] TIDAK DITEMUKAN sama sekali.")
