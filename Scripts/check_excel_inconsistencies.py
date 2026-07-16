import openpyxl

# Buka file Excel
wb = openpyxl.load_workbook(r'folder excel\Masterlist_Digitalisasi_Combined_Fix.xlsx', data_only=True)

# List sheet yang di-import
sheets_to_check = ['DL', 'CHS 2', 'CHS 3']

print("Mulai mengecek perbedaan data pada baris dengan DocumentNumber yang sama...\n")

total_issues = 0

for sheet_name in sheets_to_check:
    ws = wb[sheet_name]
    start_row = 6 if sheet_name != 'CHS 3' else 5
    
    # Kumpulkan data per DocumentNumber
    doc_data = {}
    
    for r in range(start_row, ws.max_row + 1):
        doc_val = ws.cell(row=r, column=4).value # Kolom 4 = No. Document
        if not doc_val or 'SOP/PROD/HOSE/SPS/' not in str(doc_val):
            continue
            
        doc_num = str(doc_val).strip()
        item_code = str(ws.cell(row=r, column=1).value or '').strip()
        
        # Ambil semua parameter mesin (mulai dari kolom 12 ke atas, 
        # karena kolom 1-11 biasanya Item, ID, No, DocNum, Rev, Customer, Date, dll)
        # Kita cek sampai kolom ke-60 sebagai sampel parameter
        row_params = []
        for c in range(12, 60):
            val = ws.cell(row=r, column=c).value
            row_params.append(str(val).strip() if val is not None else "")
            
        if doc_num not in doc_data:
            doc_data[doc_num] = []
            
        doc_data[doc_num].append({
            'row': r,
            'item': item_code,
            'params': row_params
        })
        
    # Sekarang bandingkan apakah ada perbedaan parameter di dalam DocumentNumber yang sama
    for doc_num, entries in doc_data.items():
        if len(entries) > 1:
            # Ambil parameter baris pertama sebagai acuan
            base_params = entries[0]['params']
            base_row = entries[0]['row']
            
            for i in range(1, len(entries)):
                curr_params = entries[i]['params']
                curr_row = entries[i]['row']
                curr_item = entries[i]['item']
                
                # Bandingkan base vs current
                diffs = []
                for col_idx, (b_val, c_val) in enumerate(zip(base_params, curr_params)):
                    if b_val != c_val:
                        # Offset + 12 karena kita potong array dari index 12
                        diffs.append(f"Kolom {col_idx + 12}: '{b_val}' vs '{c_val}'")
                
                if diffs:
                    total_issues += 1
                    print(f"⚠️ BEDA DATA pada Dokumen: {doc_num}")
                    print(f"  Baris {base_row} (Item Acuan) vs Baris {curr_row} (Item: {curr_item})")
                    for d in diffs:
                        print(f"    -> {d}")
                    print("-" * 50)

if total_issues == 0:
    print("✅ AMAN! Semua baris ganda (Item list yang banyak) memiliki nilai parameter/setting mesin yang 100% IDENTIK di file Excel.")
else:
    print(f"Ditemukan {total_issues} perbedaan data di dalam dokumen yang sama.")
