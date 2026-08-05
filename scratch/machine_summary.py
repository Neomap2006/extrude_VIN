import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

cr.execute("SELECT Machine, COUNT(*) as Jumlah FROM SpsNoDocs GROUP BY Machine ORDER BY Jumlah DESC")
rows = cr.fetchall()
print(f"{'Machine':<20} | {'Jumlah Dokumen':>15}")
print("-" * 40)
for r in rows:
    print(f"{str(r[0]):<20} | {r[1]:>15}")
