import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

# Cek semua tanggal yang ada kata "16 July" atau "16 Juli"
print("=== Cek Semua Tanggal Planning yang Mengandung '16 Jul' ===")
cur.execute("""
    SELECT DISTINCT DateShiftString 
    FROM PlanningMasters 
    WHERE DateShiftString LIKE '%16 Jul%'
""")
for r in cur.fetchall():
    print(r[0])

# Cek apakah ada item kita di sana
UPDATED_DOCS = {
    'VI-SOP-PROD-129': (2, 5), 'VI-SOP-PROD-130': (2, 6), 'VI-SOP-PROD-131': (2, 5),
    'VI-SOP-PROD-132': (2, 6), 'VI-SOP-PROD-133': (2, 6), 'VI-SOP-PROD-134': (2, 4),
    'VI-SOP-PROD-135': (2, 6), 'VI-SOP-PROD-137': (2, 7), 'VI-SOP-PROD-385': (2, 3),
    'VI-SOP-PROD-459': (2, 5), 'VI-SOP-PROD-461': (2, 3), 'SOP/ENG/25/02/12': (0, 2),
    'SOP/PROD': (0, 1), 'SOP/PROD/HOSE/24/10/054': (0, 1), 'SOP/PROD/HOSE/24/10/058': (1, 2),
    'SOP/PROD/HOSE/24/10/070': (0, 1), 'SOP/PROD/HOSE/24/10/071': (0, 1),
    'SOP/PROD/HOSE/EPR/24/10/13': (0, 1), 'SOP/PROD/HOSE/EPR/24/10/16': (0, 3),
    'SOP/PROD/HOSE/EPR/24/10/17': (0, 1), 'SOP/PROD/HOSE/EPR/24/10/22': (0, 2),
    'SOP/PROD/HOSE/EPR/24/10/24': (0, 1), 'SOP/PROD/HOSE/EPR/24/10/27': (0, 2),
    'SOP/PROD/HOSE/EPR/24/10/28': (0, 2), 'SOP/PROD/HOSE/EPR/24/10/37': (0, 1),
}
all_items = {}
for doc_num in UPDATED_DOCS.keys():
    cur.execute("SELECT ItemList, DocumentNumber FROM SpsMasters WHERE DocumentNumber = ?", (doc_num,))
    for item_list, doc in cur.fetchall():
        if item_list:
            items = [x.strip() for x in item_list.split(',') if x.strip()]
            for item in items:
                all_items[item] = doc_num

item_list_keys = list(all_items.keys())
placeholders = ','.join(['?' for _ in item_list_keys])

cur.execute(f"""
    SELECT DISTINCT p.PartName1, p.MachineName, p.DateShiftString
    FROM PlanningMasters p
    WHERE p.PartName1 IN ({placeholders}) 
    AND p.DateShiftString LIKE '%16 Jul%'
""", item_list_keys)

rows = cur.fetchall()
if rows:
    print("\nAda item kita di tanggal 16 Juli:")
    for r in rows:
        print(f"  {r[0]} di {r[1]} ({r[2]})")
else:
    print("\nTidak ada item kita di planning 16 Juli.")
conn.close()
