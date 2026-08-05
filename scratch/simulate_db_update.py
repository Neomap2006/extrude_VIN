import pyodbc
import openpyxl
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def extract_excel_mapping():
    mapping = {}
    try:
        wb = openpyxl.load_workbook(r'd:\Project\extrude\folder excel\Masterlist SPS Double Layer_Digitalisasi.xlsx', data_only=True)
        ws = wb['Parameter Setting']
        for row in ws.iter_rows(min_row=6, min_col=49, max_col=52, values_only=True):
            doc = str(row[0] or "").strip()
            mc = str(row[3] or "").strip()
            if doc and mc and doc not in ["None", "-", "nan"] and mc not in ["None", "-", "nan"]:
                mapping[doc] = mc
    except Exception as e:
        print("Error reading Digitalisasi:", e)
        
    return mapping

print("--- EXECUTING DATABASE SIMULATION ---")
mapping = extract_excel_mapping()

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

cr.execute("SELECT No, Machine, MachineCode, HoseType, DocumentNumber FROM SpsNoDocs WHERE Machine LIKE 'Non-CHS%' ORDER BY Machine, No")
rows = cr.fetchall()

rollback_sql = "-- SCRIPT UNTUK MENGEMBALIKAN DATA KE SEMULA (UNDO)\nBEGIN TRAN;\n"
update_sql = "-- SCRIPT UNTUK MENGUBAH DATA (EXECUTE)\nBEGIN TRAN;\n"

for r in rows:
    db_no = str(r[0]).strip()
    old_machine = str(r[1] or "").strip()
    old_machinecode = str(r[2] or "").strip()
    doc_num = str(r[4] or "").strip()
    
    new_machine = None
    
    if doc_num in mapping:
        new_machine = mapping[doc_num]
    elif old_machinecode and old_machinecode not in ["Non-CHS (DL)", "Non-CHS (SL)", "-"]:
        new_machine = old_machinecode
    
    if new_machine and new_machine != old_machine:
        # Quote db_no because No is an nvarchar column
        rollback_sql += f"UPDATE SpsNoDocs SET Machine = '{old_machine}' WHERE No = '{db_no}';\n"
        update_sql += f"UPDATE SpsNoDocs SET Machine = '{new_machine}' WHERE No = '{db_no}';\n"

rollback_sql += "COMMIT;\n"
update_sql += "COMMIT;\n"

with open(r"d:\Project\extrude\scratch\UNDO_Update_Machine.sql", "w") as f:
    f.write(rollback_sql)
    
with open(r"d:\Project\extrude\scratch\DO_Update_Machine.sql", "w") as f:
    f.write(update_sql)

print("Generated DO_Update_Machine.sql and UNDO_Update_Machine.sql")
