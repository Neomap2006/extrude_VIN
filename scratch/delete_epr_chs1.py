import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

# Fetch all rows that will be deleted (to create UNDO script)
cr.execute("SELECT No, DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/EPR%' AND Machine = 'CHS1'")
rows = cr.fetchall()

print(f"Baris yang akan dihapus: {len(rows)}")
for r in rows:
    print(f"  No={r[0]}, Doc={r[1]}, Machine={r[2]}")

# Build UNDO (restore) SQL using DELETE-reverse logic
# Since we're deleting, the UNDO is just a list of the deleted rows for reference
# (In real practice, to fully restore we'd need all columns - but DELETE is irreversible unless we have full data)
# So let's store key info first:

undo_note = "-- CATATAN: Baris berikut DIHAPUS dari SpsNoDocs pada update ini.\n"
undo_note += "-- Untuk mengembalikan data, Anda perlu memasukkan ulang data ini secara manual\n"
undo_note += "-- atau restore dari backup database.\n\n"
for r in rows:
    undo_note += f"-- DELETED: No={r[0]}, DocumentNumber={r[1]}, Machine={r[2]}\n"

with open(r"d:\Project\extrude\scratch\UNDO_Delete_EPR_CHS1.txt", "w") as f:
    f.write(undo_note)

# Execute DELETE
conn.autocommit = False
try:
    cr.execute("DELETE FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/EPR%' AND Machine = 'CHS1'")
    deleted = cr.rowcount
    conn.commit()
    print(f"\nBERHASIL MENGHAPUS {deleted} baris.")
    print("Catatan baris yang dihapus disimpan di: d:\\Project\\extrude\\scratch\\UNDO_Delete_EPR_CHS1.txt")
except Exception as e:
    conn.rollback()
    print("GAGAL. Rollback dijalankan. Error:", e)
