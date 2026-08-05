import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

# Cek kondisi sekarang: SPS/2 dan SPS/3 di database
print("=== SPS/2 di database ===")
cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/2/%' ORDER BY No")
sps2 = cr.fetchall()
for r in sps2:
    print(f"  No={r[0]} | {r[1]} | Machine={r[2]}")
print(f"Total SPS/2: {len(sps2)}")

print()
print("=== SPS/3 di database ===")
cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/3/%' ORDER BY No")
sps3 = cr.fetchall()
for r in sps3:
    print(f"  No={r[0]} | {r[1]} | Machine={r[2]}")
print(f"Total SPS/3: {len(sps3)}")

print()
print("=== Yang PERLU DIUPDATE ===")
sps2_wrong = [(r[0], r[1], r[2]) for r in sps2 if r[2] != 'CHS 2 Layer']
sps3_wrong = [(r[0], r[1], r[2]) for r in sps3 if r[2] != 'CHS 3 Layer']
print(f"SPS/2 yang bukan CHS 2 Layer: {len(sps2_wrong)}")
for r in sps2_wrong:
    print(f"  No={r[0]} | {r[1]} | sekarang={r[2]} -> harusnya=CHS 2 Layer")
print(f"SPS/3 yang bukan CHS 3 Layer: {len(sps3_wrong)}")
for r in sps3_wrong:
    print(f"  No={r[0]} | {r[1]} | sekarang={r[2]} -> harusnya=CHS 3 Layer")
