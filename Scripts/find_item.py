import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

item = 'TA169A'
print(f"=== Mencari Item: {item} ===")

# 1. Cari di PartMasters
cur.execute("SELECT PartCode, PartNumber, Description FROM PartMasters WHERE PartCode LIKE ?", (f'%{item}%',))
rows = cur.fetchall()
if rows:
    print(f"\n[PartMasters] Ditemukan {len(rows)} data:")
    for r in rows:
        print(f"  PartCode: {r[0]}, PartNumber: {r[1]}, Desc: {r[2]}")
else:
    print("\n[PartMasters] Tidak ditemukan.")

# 2. Cari di SpsItemLists
cur.execute("SELECT ItemList, DocumentNumber FROM SpsItemLists WHERE ItemList LIKE ?", (f'%{item}%',))
rows = cur.fetchall()
if rows:
    print(f"\n[SpsItemLists] Ditemukan {len(rows)} data:")
    for r in rows:
        print(f"  ItemList: {r[0]}, Doc: {r[1]}")
else:
    print("\n[SpsItemLists] Tidak ditemukan.")

# 3. Cari di SpsMasters
cur.execute("SELECT ItemList, DocumentNumber FROM SpsMasters WHERE ItemList LIKE ?", (f'%{item}%',))
rows = cur.fetchall()
if rows:
    print(f"\n[SpsMasters] Ditemukan {len(rows)} data:")
    for r in rows:
        print(f"  ItemList: {r[0]}, Doc: {r[1]}")
else:
    print("\n[SpsMasters] Tidak ditemukan.")

# 4. Cari di PlanningMasters
cur.execute("SELECT PartName1, PartName2, MachineName, DateShiftString FROM PlanningMasters WHERE PartName1 LIKE ?", (f'%{item}%',))
rows = cur.fetchall()
if rows:
    print(f"\n[PlanningMasters] Ditemukan {len(rows)} data:")
    for r in rows:
        print(f"  PartName1: {r[0]}, PartName2: {r[1]}, Machine: {r[2]}, Date: {r[3]}")
else:
    print("\n[PlanningMasters] Tidak ditemukan.")

conn.close()
