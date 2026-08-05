import pyodbc
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

# Mapping: DocumentNumber -> correct Machine
updates = {
    'SOP/PROD/HOSE/SPS/1/16':    'DL1',
    'SOP/PROD/HOSE/SPS/1/24':    'DL2',
    'SOP/PROD/HOSE/SPS/1/27':    'DL2',
    'SOP/PROD/HOSE/SPS/1/37':    'DL1',
    'SOP/PROD/HOSE/SPS/1/38':    'DL2',
    'SOP/PROD/HOSE/SPS/24/10/16':'DL1',
    'SOP/PROD/HOSE/SPS/24/10/17':'DL2',
    'SOP/PROD/HOSE/SPS/24/10/24':'DL2',
    'SOP/PROD/HOSE/SPS/24/10/27':'DL2',
    'SOP/PROD/HOSE/SPS/24/10/37':'DL1',
    'SOP/PROD/HOSE/SPS/24/12/38':'DL3/2',
    'SOP/PROD/HOSE/SPS/3/16':    'CHS 3 Layer',
}

# UNDO script
undo_sql = "-- UNDO: Kembalikan CHS1 yang sudah diupdate\nBEGIN TRAN;\n"
for doc in updates.keys():
    undo_sql += f"UPDATE SpsNoDocs SET Machine = 'CHS1' WHERE DocumentNumber = '{doc}';\n"
undo_sql += "COMMIT;\n"

with open(r"d:\Project\extrude\scratch\UNDO_Fix_CHS1.sql", "w") as f:
    f.write(undo_sql)

# Execute updates
conn.autocommit = False
total = 0
try:
    for doc, new_machine in updates.items():
        cr.execute(f"UPDATE SpsNoDocs SET Machine = ? WHERE DocumentNumber = ? AND Machine = 'CHS1'", new_machine, doc)
        affected = cr.rowcount
        total += affected
        print(f"  {doc} -> {new_machine}  ({affected} baris)")
    conn.commit()
    print(f"\nBERHASIL UPDATE {total} baris total.")
    print(f"File UNDO tersimpan di: d:\\Project\\extrude\\scratch\\UNDO_Fix_CHS1.sql")
except Exception as e:
    conn.rollback()
    print("GAGAL. Rollback dijalankan. Error:", e)
