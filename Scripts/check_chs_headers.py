import pyodbc
import openpyxl
import sys
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook(r'folder excel\Masterlist_Digitalisasi_Combined_Fix.xlsx', data_only=True)

for sheet_name in ['CHS 2', 'CHS 3']:
    ws = wb[sheet_name]
    print(f"\n=== Sheet: {sheet_name} ===")
    # Cek baris 1-6 kolom 1-10 untuk cari header/machine code
    for r in range(1, 8):
        row_vals = []
        for c in range(1, 12):
            v = ws.cell(row=r, column=c).value
            row_vals.append(f"col{c}=[{str(v)[:20] if v else ''}]")
        print(f"  Row {r}: {' | '.join(row_vals)}")
    
    # Cek beberapa data row
    print(f"\n  --- Data rows ---")
    for r in range(7, 12):
        row_vals = []
        for c in range(1, 12):
            v = ws.cell(row=r, column=c).value
            row_vals.append(f"col{c}=[{str(v)[:15] if v else ''}]")
        print(f"  Row {r}: {' | '.join(row_vals)}")
