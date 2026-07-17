import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

UPDATED_DOCS = {
    'VI-SOP-PROD-129': (2, 5),
    'VI-SOP-PROD-130': (2, 6),
    'VI-SOP-PROD-131': (2, 5),
    'VI-SOP-PROD-132': (2, 6),
    'VI-SOP-PROD-133': (2, 6),
    'VI-SOP-PROD-134': (2, 4),
    'VI-SOP-PROD-135': (2, 6),
    'VI-SOP-PROD-137': (2, 7),
    'VI-SOP-PROD-385': (2, 3),
    'VI-SOP-PROD-459': (2, 5),
    'VI-SOP-PROD-461': (2, 3),
    'SOP/ENG/25/02/12': (0, 2),
    'SOP/PROD': (0, 1),
    'SOP/PROD/HOSE/24/10/054': (0, 1),
    'SOP/PROD/HOSE/24/10/058': (1, 2),
    'SOP/PROD/HOSE/24/10/070': (0, 1),
    'SOP/PROD/HOSE/24/10/071': (0, 1),
    'SOP/PROD/HOSE/EPR/24/10/13': (0, 1),
    'SOP/PROD/HOSE/EPR/24/10/16': (0, 3),
    'SOP/PROD/HOSE/EPR/24/10/17': (0, 1),
    'SOP/PROD/HOSE/EPR/24/10/22': (0, 2),
    'SOP/PROD/HOSE/EPR/24/10/24': (0, 1),
    'SOP/PROD/HOSE/EPR/24/10/27': (0, 2),
    'SOP/PROD/HOSE/EPR/24/10/28': (0, 2),
    'SOP/PROD/HOSE/EPR/24/10/37': (0, 1),
}

# Step 1: Kumpulkan semua ItemList dari SpsMasters untuk dokumen yang diupdate
all_items = {}  # item_code -> doc_number
for doc_num in UPDATED_DOCS.keys():
    cur.execute("SELECT ItemList, DocumentNumber FROM SpsMasters WHERE DocumentNumber = ?", (doc_num,))
    for item_list, doc in cur.fetchall():
        if item_list:
            items = [x.strip() for x in item_list.split(',') if x.strip()]
            for item in items:
                all_items[item] = doc_num

print(f"Total item dari dokumen yang diupdate: {len(all_items)}")
print()

# Step 2: Cari item-item ini di PlanningMasters
if all_items:
    item_list_keys = list(all_items.keys())
    placeholders = ','.join(['?' for _ in item_list_keys])
    
    cur.execute(f"""
        SELECT DISTINCT p.PartName1, p.PartName2, p.MachineName, p.DateShiftString, p.CreatedAt
        FROM PlanningMasters p
        WHERE p.PartName1 IN ({placeholders})
        ORDER BY p.DateShiftString DESC
    """, item_list_keys)
    
    planning_rows = cur.fetchall()
    
    if planning_rows:
        print(f"=== ITEM PLANNING YANG MENGGUNAKAN DOKUMEN YANG DIUPDATE HARI INI ===")
        print(f"(Hanya READ ONLY - Tidak ada perubahan data)\n")
        print(f"{'No':<4} {'Kode Item':<14} {'No Dokumen SPS':<30} {'Rev Lama':<10} {'Rev Baru':<10} {'Mesin':<35} {'Tanggal Shift'}")
        print("-" * 145)
        
        for i, (part1, part2, machine, date_shift, created) in enumerate(planning_rows, 1):
            doc = all_items.get(part1, '-')
            old_rev, new_rev = UPDATED_DOCS.get(doc, ('-', '-'))
            print(f"{i:<4} {part1:<14} {doc:<30} Rev {old_rev:<6} Rev {new_rev:<6} {machine:<35} {date_shift}")
        
        print()
        print(f"Total: {len(planning_rows)} item planning yang terdampak perubahan revisi hari ini.")
        
    else:
        print("Tidak ada item dari planning yang menggunakan dokumen yang diupdate hari ini.")
        print("(Semua item di planning menggunakan dokumen SPS yang berbeda atau tidak terdaftar)")

conn.close()
