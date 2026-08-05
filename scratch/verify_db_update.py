import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

cr.execute("SELECT COUNT(*) FROM SpsNoDocs WHERE Machine LIKE 'Non-CHS%'")
remaining = cr.fetchone()[0]
print('Remaining Non-CHS rows:', remaining)

cr.execute("SELECT TOP 10 No, DocumentNumber, Machine FROM SpsNoDocs WHERE Machine LIKE 'DL%'")
rows = cr.fetchall()
print('Sample DL records now:')
for r in rows:
    print(r)
