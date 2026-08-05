import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE Machine = 'CHS1' ORDER BY DocumentNumber")
rows = cr.fetchall()
print(f"Total CHS1 yang tersisa: {len(rows)}")
print()
for r in rows:
    print(f"No={r[0]}  |  {r[1]}")
