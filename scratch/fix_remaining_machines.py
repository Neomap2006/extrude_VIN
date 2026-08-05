import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

updates = [
    ('DL 1',          'DL1'),
    ('Double Layer',  'DL1/2'),
    ('Non-CHS (DL)',  'DL1'),
]

conn.autocommit = False
try:
    total = 0
    for old, new in updates:
        cr.execute("UPDATE SpsNoDocs SET Machine = ? WHERE Machine = ?", new, old)
        affected = cr.rowcount
        total += affected
        print(f"  '{old}' -> '{new}' : {affected} baris")
    conn.commit()
    print(f"\nBERHASIL! Total {total} baris diupdate.")
except Exception as e:
    conn.rollback()
    print("GAGAL. Error:", e)

print()
print("=== REKAP MESIN FINAL ===")
cr.execute("SELECT Machine, COUNT(*) FROM SpsNoDocs GROUP BY Machine ORDER BY COUNT(*) DESC")
for r in cr.fetchall():
    print(f"  {str(r[0]):<20} : {r[1]} dokumen")
