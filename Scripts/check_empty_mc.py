import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

# Cek 7 dokumen yang masih kosong
cur.execute("""
    SELECT DocumentNumber, MachineCode, HoseType, InnerTube, OuterCover
    FROM SpsNoDocs
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%'
    AND MachineCode = ''
    ORDER BY DocumentNumber
""")
rows = cur.fetchall()
print(f"Dokumen MachineCode = '' ({len(rows)} records):")
for r in rows:
    print(f"  {r[0]} | machine=[{r[1]}] | hose={r[2]} | inner={r[3]} | outer={r[4]}")
