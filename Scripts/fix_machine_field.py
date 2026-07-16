import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

# Cek dulu kondisi kolom Machine vs MachineCode untuk SPS/2 dan SPS/3
cur.execute("""
    SELECT DocumentNumber, Machine, MachineCode
    FROM SpsNoDocs
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/2/%'
    OR DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/3/%'
    OR DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/1/%'
    OR (DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/24/%' AND (Machine IS NULL OR Machine = ''))
    ORDER BY DocumentNumber
""")
rows = cur.fetchall()
print(f"{'DocumentNumber':<35} {'Machine':<15} {'MachineCode':<15}")
print("-"*65)
for r in rows[:15]:
    print(f"{r[0]:<35} [{str(r[1] or ''):<13}] [{str(r[2] or ''):<13}]")
print(f"... total {len(rows)} records")

# Update Machine field sesuai MachineCode
# CHS2
cur.execute("""
    UPDATE SpsNoDocs SET Machine = 'CHS2'
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/2/%'
    AND (Machine IS NULL OR Machine = '')
""")
print(f"\nUpdate Machine=CHS2: {cur.rowcount} records")

# CHS3
cur.execute("""
    UPDATE SpsNoDocs SET Machine = 'CHS3'
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/3/%'
    AND (Machine IS NULL OR Machine = '')
""")
print(f"Update Machine=CHS3: {cur.rowcount} records")

# CHS1 (SPS/1/* dan SPS/24/*)
cur.execute("""
    UPDATE SpsNoDocs SET Machine = 'CHS1'
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%'
    AND (Machine IS NULL OR Machine = '')
    AND MachineCode = 'CHS1'
""")
print(f"Update Machine=CHS1: {cur.rowcount} records")

conn.commit()

# Verify
print("\n=== Verifikasi akhir ===")
cur.execute("""
    SELECT Machine, MachineCode, COUNT(*) as cnt
    FROM SpsNoDocs
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%'
    GROUP BY Machine, MachineCode
    ORDER BY Machine, MachineCode
""")
for row in cur.fetchall():
    print(f"  Machine=[{row[0]}] | MachineCode=[{row[1]}] -> {row[2]} dokumen")
