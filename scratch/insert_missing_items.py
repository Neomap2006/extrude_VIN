import pandas as pd
import pyodbc

# Part yang dicari
missing_parts = ['HN0490', 'HN0590', 'HN0540', 'HN0530', 'PM1110', 'HN0300', 'HN0550', 'HN0250', 'HN0560']

# 1. Baca Excel EXPORT_PIVOT
file_path = r"d:\Project\extrude\folder excel\EXPORT_PIVOT.xlsx"
df = pd.read_excel(file_path, sheet_name="Sum", header=None)

# Cari header
header_row_idx = 0
for i in range(10):
    row_vals = df.iloc[i].astype(str).str.upper().tolist()
    if any('VIN' in v for v in row_vals) or any('SHRINK' in v for v in row_vals):
        header_row_idx = i
        break

df.columns = df.iloc[header_row_idx]
df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
df.columns = [str(c).strip() for c in df.columns]

part_col = [c for c in df.columns if 'VIN' in c.upper() or 'PART NO' in c.upper()][0]
sps_col = [c for c in df.columns if 'NO SPS' in c.upper() or 'SOP/PROD' in c.upper() or 'NO. DOC' in c.upper() or 'DOC' in c.upper()]
shrink_col = [c for c in df.columns if 'SHRINK' in c.upper()][0]

if not sps_col:
    # If standard names fail, let's just pick the column that has "SOP/PROD" data in it for the first few rows
    for c in df.columns:
        vals = df[c].astype(str).str.upper().tolist()
        if any('SOP/PROD' in str(v) for v in vals[:10]):
            sps_col = c
            break
else:
    sps_col = sps_col[0]

print(f"Part column: {part_col}")
print(f"SPS Document column: {sps_col}")
print(f"Shrink column: {shrink_col}")

# Parsing helper (dari script sebelumnya)
def parse_shrink(raw):
    raw = str(raw).strip().replace(',', '.')
    if raw in ('nan', '', '-', 'N/A'): return None, None, None
    if '±' in raw or '+/-' in raw.lower():
        raw2 = raw.replace('+/-', '±')
        parts = raw2.split('±')
        try:
            std = float(parts[0].strip())
            tol = float(parts[1].strip())
            return round(std - tol, 2), round(std, 2), round(std + tol, 2)
        except: pass
    if '-' in raw:
        parts = raw.split('-')
        if len(parts) == 2:
            try:
                p1, p2 = float(parts[0].strip()), float(parts[1].strip())
                mn, mx = min(p1, p2), max(p1, p2)
                return mn, round((mn + mx) / 2, 2), mx
            except: pass
    try:
        v = float(raw)
        return (None, v, None) if v == 0 else (v, v, v)
    except: return None, None, None

# Kumpulkan data dari excel
items_to_add = []
for part in missing_parts:
    row = df[df[part_col].astype(str).str.strip().str.upper() == part.upper()]
    if not row.empty:
        sps_no = str(row[sps_col].iloc[0]).strip()
        raw_shrink = row[shrink_col].iloc[0]
        mn, std, mx = parse_shrink(raw_shrink)
        items_to_add.append({
            'ItemList': part.upper(),
            'DocumentNumber': sps_no,
            'Min': mn,
            'Std': std,
            'Max': mx
        })

# 2. Cek database
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(
    f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;'
    f'UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes'
)
cr = conn.cursor()

added = 0
for item in items_to_add:
    doc = item['DocumentNumber']
    part = item['ItemList']
    
    if not doc or doc.lower() == 'nan':
        print(f"[{part}] No Document kosong di Excel.")
        continue
        
    # Cek apakah No Document ada di tabel SpsNoDocs
    cr.execute("SELECT DocumentNumber FROM SpsNoDocs WHERE DocumentNumber = ?", doc)
    doc_exists = cr.fetchone()
    
    if doc_exists:
        # Tambahkan ke SpsItemLists
        cr.execute("""
            INSERT INTO SpsItemLists (ItemList, DocumentNumber, LengthShrinkMin, LengthShrinkStd, LengthShrinkMax)
            VALUES (?, ?, ?, ?, ?)
        """, part, doc, item['Min'], item['Std'], item['Max'])
        print(f"[{part}] BERHASIL DITAMBAHKAN -> Document: {doc}")
        added += 1
    else:
        print(f"[{part}] GAGAL -> Document '{doc}' ADA di Excel tapi TIDAK ADA di Master Database (SpsNoDocs).")

conn.commit()
conn.close()

print(f"\nProses selesai. Berhasil menambahkan {added} item list.")
