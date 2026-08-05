import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()
conn.autocommit = False

cr.execute("UPDATE SpsNoDocs SET Machine = 'CHS 2 Layer' WHERE DocumentNumber IN ('VI-SOP-PROD-136', 'VI-SOP-PROD-460')")
updated = cr.rowcount
conn.commit()
print(f'BERHASIL! {updated} baris diupdate ke CHS 2 Layer.')

print()
print('=== REKAP MESIN FINAL ===')
cr.execute('SELECT Machine, COUNT(*) FROM SpsNoDocs GROUP BY Machine ORDER BY COUNT(*) DESC')
for r in cr.fetchall():
    print(f'  {str(r[0]):<20} : {r[1]} dokumen')
