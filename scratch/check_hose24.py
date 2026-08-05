import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

# Cek dokumen format SOP/PROD/HOSE/24/10/0xx (tanpa SPS/)
cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/24/%' AND DocumentNumber NOT LIKE '%/SPS/%' AND DocumentNumber NOT LIKE '%/EPR/%' ORDER BY DocumentNumber")
rows = cr.fetchall()
print(f"Total dokumen format HOSE/24/... (tanpa SPS/EPR): {len(rows)}")
for r in rows:
    print(f"  No={r[0]} | {r[1]} | Machine={r[2]}")
