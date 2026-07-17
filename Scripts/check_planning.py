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

doc_list = list(UPDATED_DOCS.keys())
placeholders = ','.join(['?' for _ in doc_list])

# Cek kolom di SpsNoDocs yang bisa berisi ItemList
print("=== Item List dari SpsNoDocs ===")
# Cek apakah ada kolom ItemList di SpsNoDocs
cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpsNoDocs' AND COLUMN_NAME LIKE '%Item%'")
item_cols = [row[0] for row in cur.fetchall()]
print("Kolom dengan 'Item' di SpsNoDocs:", item_cols)

# Ambil data ItemList dari SpsNoDocs untuk dokumen yang diupdate
print()
print("=== Mengambil data dari SpsNoDocs untuk dokumen yang diupdate ===")
all_items = {}  # item_code -> doc_number

for doc_num in doc_list:
    # Cek di SpsMasters - ada ItemList per baris
    cur.execute("SELECT ItemList, DocumentNumber FROM SpsMasters WHERE DocumentNumber = ?", (doc_num,))
    rows = cur.fetchall()
    if rows:
        for item_list, doc in rows:
            if item_list:
                items = [x.strip() for x in item_list.split(',') if x.strip()]
                for item in items:
                    all_items[item] = doc_num

if all_items:
    print(f"Ditemukan {len(all_items)} item dari SpsMasters")
    for item, doc in list(all_items.items())[:10]:
        print(f"  {item} -> {doc}")
else:
    print("Tidak ada item di SpsMasters untuk dokumen yang diupdate")

# Also check SpsNoDocs directly - maybe it has a different column
print()
print("=== Semua kolom SpsNoDocs ===")
cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='SpsNoDocs' ORDER BY ORDINAL_POSITION")
all_cols = [row[0] for row in cur.fetchall()]
print(all_cols)

# Sample SpsNoDocs for one of the updated docs
print()
print("=== Sample SpsNoDocs untuk VI-SOP-PROD-129 (kolom awal) ===")
cur.execute("SELECT No, Machine, DocumentNumber, RevisionNumber, Customer, HoseType FROM SpsNoDocs WHERE DocumentNumber = 'VI-SOP-PROD-129'")
for row in cur.fetchall():
    print(f"  {row}")

conn.close()
