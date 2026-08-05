import openpyxl
import pandas as pd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("=== Cek file: Masterlist_Digitalisasi_Combined_Fix.xlsx (sheet: DL) ===")
try:
    df = pd.read_excel(r'd:\Project\extrude\folder excel\Masterlist_Digitalisasi_Combined_Fix.xlsx', sheet_name='DL', header=None)
    header_idx = -1
    for i in range(15):
        row_str = ' '.join([str(x).upper() for x in df.iloc[i].values])
        if 'MC' in row_str:
            header_idx = i
            break
    df.columns = [str(c).strip().upper() for c in df.iloc[header_idx]]
    df = df.iloc[header_idx+1:]
    doc_cols = [c for c in df.columns if 'DOCUMENT' in c or 'DOC' in c]
    mc_cols = [c for c in df.columns if c == 'MC']
    
    found_sps2 = []
    for _, row in df.iterrows():
        doc = str(row[doc_cols[0]]).strip()
        mc = str(row[mc_cols[0]]).strip()
        if 'SPS/2/' in doc and doc != 'nan':
            found_sps2.append((doc, mc))
            print(f"  {doc} -> {mc}")
    print(f"Total SPS/2 di Combined DL Excel: {len(found_sps2)}")
except Exception as e:
    print("Error Combined:", e)

print()
print("=== Cek file: Masterlist SPS Double Layer_Digitalisasi.xlsx ===")
try:
    wb = openpyxl.load_workbook(r'd:\Project\extrude\folder excel\Masterlist SPS Double Layer_Digitalisasi.xlsx', data_only=True)
    ws = wb['Parameter Setting']
    count = 0
    for row in ws.iter_rows(min_row=6, min_col=49, max_col=52, values_only=True):
        doc = str(row[0] or '').strip()
        mc = str(row[3] or '').strip()
        if 'SPS/2/' in doc:
            count += 1
            print(f"  {doc} -> {mc}")
    print(f"Total SPS/2 di DL Digitalisasi: {count}")
except Exception as e:
    print("Error DL Digitalisasi:", e)

print()
print("=== Cek file Excel CHS 2 Layer (cari sheet yg relevan) ===")
import os
excel_dir = r'd:\Project\extrude\folder excel'
for fname in os.listdir(excel_dir):
    if 'CHS' in fname.upper() or 'SINGLE' in fname.upper():
        print(f"  File ditemukan: {fname}")
