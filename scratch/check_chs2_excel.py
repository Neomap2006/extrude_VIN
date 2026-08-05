import openpyxl
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Check which SPS/2 documents exist in the CHS 2 Layer Excel
wb = openpyxl.load_workbook(r'd:\Project\extrude\folder excel\Masterlist SPS CHS 2 Layer DIG.xlsx', data_only=True)
ws = wb['Parameter Setting']

print("=== SPS/2 documents in CHS 2 Layer Excel ===")
found = []
for row in ws.iter_rows(min_row=7, values_only=True):
    # Col 4 (index 3) = No. Document
    doc = str(row[3] or '').strip()
    machine = str(row[2] or '').strip()
    if 'SPS/2/' in doc:
        found.append((doc, machine))
        print(f"  {doc} -> {machine}")

print(f"\nTotal ditemukan: {len(found)}")

# Also check what documents SPS/2 are in DB with Machine = DL1 or DL2
import pyodbc
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

print()
print("=== SPS/2 di database (semua mesin) ===")
cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/2/%' ORDER BY No, DocumentNumber")
rows = cr.fetchall()
print(f"Total: {len(rows)} baris")
for r in rows:
    print(f"  No={r[0]} | {r[1]} | Machine={r[2]}")
