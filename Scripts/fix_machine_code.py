import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

# CHS 2: DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/2/%'
# CHS 3: DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/3/%'
# CHS 1 (sheet DL for single layer): DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/1/%'
# SPS/24/... → juga CHS1 (single layer) dari sheet DL

# Dari check sebelumnya:
# SPS/1/* dan SPS/24/* → MachineCode = NULL/empty → ini sheet CHS 1 (atau DL?)
# Tapi DL1/DL2/DL3 sudah benar
# Yang NULL adalah SPS/2/* dan SPS/3/*

print("=== SEBELUM UPDATE ===")
cur.execute("""
    SELECT DocumentNumber, MachineCode 
    FROM SpsNoDocs 
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%' 
    AND (MachineCode IS NULL OR MachineCode = '')
    ORDER BY DocumentNumber
""")
rows = cur.fetchall()
print(f"Dokumen dengan MachineCode kosong: {len(rows)}")
for r in rows[:10]:
    print(f"  {r[0]} -> [{r[1]}]")
if len(rows) > 10:
    print(f"  ... dan {len(rows)-10} lainnya")

# Update CHS 2
cur.execute("""
    UPDATE SpsNoDocs SET MachineCode = 'CHS2'
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/2/%'
    AND (MachineCode IS NULL OR MachineCode = '')
""")
cnt2 = cur.rowcount
print(f"\nUpdate CHS2: {cnt2} records")

# Update CHS 3
cur.execute("""
    UPDATE SpsNoDocs SET MachineCode = 'CHS3'
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/3/%'
    AND (MachineCode IS NULL OR MachineCode = '')
""")
cnt3 = cur.rowcount
print(f"Update CHS3: {cnt3} records")

# Update CHS 1 (SPS/1/* yang kosong)
cur.execute("""
    UPDATE SpsNoDocs SET MachineCode = 'CHS1'
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/1/%'
    AND (MachineCode IS NULL OR MachineCode = '')
""")
cnt1 = cur.rowcount
print(f"Update CHS1: {cnt1} records")

conn.commit()

print("\n=== SESUDAH UPDATE ===")
cur.execute("""
    SELECT MachineCode, COUNT(*) as cnt 
    FROM SpsNoDocs 
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%' 
    GROUP BY MachineCode 
    ORDER BY MachineCode
""")
for row in cur.fetchall():
    print(f"  [{row[0]}] -> {row[1]} dokumen")
