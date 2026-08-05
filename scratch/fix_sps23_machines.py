import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

# Ambil data lama dulu untuk UNDO
cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/2/%' AND Machine != 'CHS 2 Layer'")
sps2_old = cr.fetchall()
cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/3/%' AND Machine != 'CHS 3 Layer'")
sps3_old = cr.fetchall()

# Generate UNDO script
undo = "-- UNDO: Kembalikan Machine SPS/2 dan SPS/3 ke nilai lama\nBEGIN TRAN;\n"
for r in sps2_old:
    undo += f"UPDATE SpsNoDocs SET Machine = '{r[2]}' WHERE DocumentNumber = '{r[1]}';\n"
for r in sps3_old:
    undo += f"UPDATE SpsNoDocs SET Machine = '{r[2]}' WHERE DocumentNumber = '{r[1]}';\n"
undo += "COMMIT;\n"

with open(r"d:\Project\extrude\scratch\UNDO_Fix_SPS23.sql", "w") as f:
    f.write(undo)

# Execute update
conn.autocommit = False
try:
    cr.execute("UPDATE SpsNoDocs SET Machine = 'CHS 2 Layer' WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/2/%' AND Machine != 'CHS 2 Layer'")
    updated_sps2 = cr.rowcount

    cr.execute("UPDATE SpsNoDocs SET Machine = 'CHS 3 Layer' WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/3/%' AND Machine != 'CHS 3 Layer'")
    updated_sps3 = cr.rowcount

    conn.commit()
    print(f"BERHASIL!")
    print(f"  SPS/2 diupdate ke CHS 2 Layer : {updated_sps2} baris")
    print(f"  SPS/3 diupdate ke CHS 3 Layer : {updated_sps3} baris")
    print(f"  Total                          : {updated_sps2 + updated_sps3} baris")
    print(f"\nFile UNDO tersimpan di: d:\\Project\\extrude\\scratch\\UNDO_Fix_SPS23.sql")
except Exception as e:
    conn.rollback()
    print("GAGAL. Rollback dijalankan. Error:", e)

# Tampilkan rekapitulasi akhir
print()
print("=== REKAP MESIN TERBARU DI DATABASE ===")
cr.execute("SELECT Machine, COUNT(*) as Jumlah FROM SpsNoDocs GROUP BY Machine ORDER BY Jumlah DESC")
for r in cr.fetchall():
    print(f"  {str(r[0]):<20} : {r[1]} dokumen")
