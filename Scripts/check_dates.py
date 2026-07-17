import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

cur.execute("""
    SELECT DISTINCT DateShiftString 
    FROM PlanningMasters 
""")
print("=== Semua Tanggal di PlanningMasters ===")
for r in cur.fetchall():
    print(r[0])

conn.close()
