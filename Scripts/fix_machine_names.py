import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

# Fix CHS2 -> CHS 2 Layer
cur.execute("""
    UPDATE SpsNoDocs 
    SET Machine = 'CHS 2 Layer', MachineCode = 'CHS 2 Layer'
    WHERE MachineCode = 'CHS2' OR Machine = 'CHS2'
""")
cnt2 = cur.rowcount
print(f"Updated CHS2 to 'CHS 2 Layer': {cnt2} records")

# Fix CHS3 -> CHS 3 Layer
cur.execute("""
    UPDATE SpsNoDocs 
    SET Machine = 'CHS 3 Layer', MachineCode = 'CHS 3 Layer'
    WHERE MachineCode = 'CHS3' OR Machine = 'CHS3'
""")
cnt3 = cur.rowcount
print(f"Updated CHS3 to 'CHS 3 Layer': {cnt3} records")

conn.commit()
print("Done!")
