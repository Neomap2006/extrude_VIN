import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

cur.execute("SELECT MachineCode, COUNT(*) as cnt FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%' GROUP BY MachineCode ORDER BY MachineCode")
print("MachineCode distribution:")
for row in cur.fetchall():
    print(f"  [{row[0]}] -> {row[1]} dokumen")
