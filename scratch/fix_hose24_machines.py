import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

# Build list of docs to update
chs2_docs = [f'SOP/PROD/HOSE/24/10/0{n}' for n in range(38, 56)]   # 038-055
chs3_docs = [f'SOP/PROD/HOSE/24/10/0{n}' for n in range(56, 72)]   # 056-071

# UNDO: save old values first
cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/24/10/0%'")
old_rows = cr.fetchall()
undo = "-- UNDO: kembalikan mesin dokumen HOSE/24\nBEGIN TRAN;\n"
for r in old_rows:
    undo += f"UPDATE SpsNoDocs SET Machine = '{r[2]}' WHERE DocumentNumber = '{r[1]}';\n"
undo += "COMMIT;\n"
with open(r"d:\Project\extrude\scratch\UNDO_Fix_HOSE24.sql", "w") as f:
    f.write(undo)

# Execute updates
conn.autocommit = False
try:
    # CHS 2 Layer: 038-055
    placeholders = ','.join(['?' for _ in chs2_docs])
    cr.execute(f"UPDATE SpsNoDocs SET Machine = 'CHS 2 Layer' WHERE DocumentNumber IN ({placeholders}) AND Machine != 'CHS 2 Layer'", *chs2_docs)
    updated_chs2 = cr.rowcount

    # CHS 3 Layer: 056-071
    placeholders = ','.join(['?' for _ in chs3_docs])
    cr.execute(f"UPDATE SpsNoDocs SET Machine = 'CHS 3 Layer' WHERE DocumentNumber IN ({placeholders}) AND Machine != 'CHS 3 Layer'", *chs3_docs)
    updated_chs3 = cr.rowcount

    conn.commit()
    print("BERHASIL!")
    print(f"  038-055 -> CHS 2 Layer : {updated_chs2} baris diupdate")
    print(f"  056-071 -> CHS 3 Layer : {updated_chs3} baris diupdate")
    print(f"  Total                  : {updated_chs2 + updated_chs3} baris")
    print(f"\nFile UNDO: d:\\Project\\extrude\\scratch\\UNDO_Fix_HOSE24.sql")
except Exception as e:
    conn.rollback()
    print("GAGAL. Rollback. Error:", e)

# Rekap akhir
print()
print("=== REKAP MESIN FINAL ===")
cr.execute("SELECT Machine, COUNT(*) FROM SpsNoDocs GROUP BY Machine ORDER BY COUNT(*) DESC")
for r in cr.fetchall():
    print(f"  {str(r[0]):<20} : {r[1]} dokumen")
