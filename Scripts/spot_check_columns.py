import pyodbc

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

# Spot check: ambil 1 dokumen dari tiap sheet (CHS1, CHS2, CHS3, DL)
test_docs = [
    'SOP/PROD/HOSE/SPS/1/1',   # CHS 1 (format X-Y)
    'SOP/PROD/HOSE/SPS/2/1',   # CHS 2 (format ±)
    'SOP/PROD/HOSE/SPS/3/1',   # CHS 3 (format ±, 3 layer)
]

# Kolom yang ingin dicek
check_cols = [
    ('Nipple',        'Nipple_Min',        'Nipple_Asli',        'Nipple_Max'),
    ('TubeDie',       'TubeDie_Min',       'TubeDie_Asli',       'TubeDie_Max'),
    ('HeadTemp1',     'HeadTemp1_Min',     'HeadTemp1_Asli',     'HeadTemp1_Max'),
    ('ScrewSpeed1',   'ScrewSpeed1_Min',   'ScrewSpeed1_Asli',   'ScrewSpeed1_Max'),
    ('Pressure1',     'Pressure1_Min',     'Pressure1_Asli',     'Pressure1_Max'),
    ('FeedRollRatio1','FeedRollRatio1_Min','FeedRollRatio1_Asli','FeedRollRatio1_Max'),
    ('ChillerWaterTemp','ChillerWaterTemp_Min','ChillerWaterTemp_Asli','ChillerWaterTemp_Max'),
]

select_cols = ['DocumentNumber']
for group in check_cols:
    select_cols.extend(group)

cur.execute(f"SELECT {', '.join(select_cols)} FROM SpsNoDocs WHERE DocumentNumber IN ({','.join(['?']*len(test_docs))})", test_docs)
rows = cur.fetchall()

print(f"\n{'='*110}")
print(f"{'Column':<22} {'Raw Text':<16} {'MIN':>10} {'STN (Asli)':>12} {'MAX':>10}  Status")
print(f"{'='*110}")

for row in rows:
    doc = row[0]
    print(f"\n--- {doc} ---")
    for i, (raw_col, min_col, asli_col, max_col) in enumerate(check_cols):
        offset = 1 + i * 4
        raw   = str(row[offset] or '').strip()
        mn    = row[offset+1]
        asli  = row[offset+2]
        mx    = row[offset+3]
        
        if not raw:
            status = '(kosong di Excel)'
        elif mn is not None and asli is not None and mx is not None:
            status = 'OK - MIN+STN+MAX'
        elif mn is not None and mx is not None and asli is None:
            status = 'OK - MIN+MAX (range)'
        elif mx is not None and mn is None and asli is None:
            status = 'OK - MAX only (Max. format)'
        elif mn is not None or asli is not None or mx is not None:
            status = 'WARN - Parsial'
        else:
            status = 'GAGAL - Tidak terparse!'
            
        print(f"  {raw_col:<20} {raw:<16} {str(mn or ''):>10} {str(asli or ''):>12} {str(mx or ''):>10}  {status}")
