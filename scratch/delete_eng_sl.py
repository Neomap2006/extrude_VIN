import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

target_docs = ['SOP/ENG/25/02/14', 'SOP/ENG/26/04/10']

# Cek dulu sebelum hapus
cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber IN ('SOP/ENG/25/02/14', 'SOP/ENG/26/04/10')")
rows = cr.fetchall()
print(f"Baris yang akan dihapus: {len(rows)}")
for r in rows:
    print(f"  No={r[0]} | {r[1]} | {r[2]}")

# UNDO note
undo = "-- Baris yang dihapus (untuk referensi):\n"
for r in rows:
    undo += f"-- No={r[0]}, DocumentNumber={r[1]}, Machine={r[2]}\n"
with open(r"d:\Project\extrude\scratch\UNDO_Delete_ENG_SL.txt", "w") as f:
    f.write(undo)

# Execute DELETE
conn.autocommit = False
try:
    cr.execute("DELETE FROM SpsNoDocs WHERE DocumentNumber IN ('SOP/ENG/25/02/14', 'SOP/ENG/26/04/10')")
    deleted = cr.rowcount
    conn.commit()
    print(f"\nBERHASIL MENGHAPUS {deleted} baris.")
except Exception as e:
    conn.rollback()
    print("GAGAL. Error:", e)

# Tampilkan rekap akhir
print()
print("=== REKAP MESIN SEKARANG ===")
cr.execute("SELECT Machine, COUNT(*) FROM SpsNoDocs GROUP BY Machine ORDER BY COUNT(*) DESC")
for r in cr.fetchall():
    print(f"  {str(r[0]):<20} : {r[1]} dokumen")
