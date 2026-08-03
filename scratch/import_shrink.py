"""
Import Length Shrink (mm) dari EXPORT_PIVOT.xlsx sheet 'Sum'
ke tabel SpsItemLists di SQL Server.
Script ini hanya UPDATE kolom LengthShrinkMin/Std/Max.
Tidak ada data yang dihapus.
"""
import pyodbc
import pandas as pd

# =====================
# 1. Baca File Excel
# =====================
file_path = r"d:\Project\extrude\folder excel\EXPORT_PIVOT.xlsx"
df = pd.read_excel(file_path, sheet_name="Sum", header=None)

# Cari header row (baris yang mengandung 'VIN' atau 'SHRINK')
header_row_idx = 0
for i in range(10):
    row_vals = df.iloc[i].astype(str).str.upper().tolist()
    if any('VIN' in v for v in row_vals) or any('SHRINK' in v for v in row_vals):
        header_row_idx = i
        break

df.columns = df.iloc[header_row_idx]
df = df.iloc[header_row_idx + 1:].reset_index(drop=True)

# Bersihkan nama kolom
df.columns = [str(c).strip() for c in df.columns]
part_col   = [c for c in df.columns if 'VIN' in c.upper() or 'PART NO' in c.upper()][0]
shrink_col = [c for c in df.columns if 'SHRINK' in c.upper()][0]

print(f"Part column   : {part_col}")
print(f"Shrink column : {shrink_col}")

# Parsing helper
def parse_shrink(raw):
    raw = str(raw).strip().replace(',', '.')
    if raw in ('nan', '', '-', 'N/A'):
        return None, None, None
    # Format "X±Y"
    if '±' in raw or '+/-' in raw.lower():
        raw2 = raw.replace('+/-', '±')
        parts = raw2.split('±')
        try:
            std = float(parts[0].strip())
            tol = float(parts[1].strip())
            return round(std - tol, 2), round(std, 2), round(std + tol, 2)
        except:
            pass
    # Format "X-Y" (rentang, bukan negatif)
    if '-' in raw:
        parts = raw.split('-')
        if len(parts) == 2:
            try:
                p1, p2 = float(parts[0].strip()), float(parts[1].strip())
                mn, mx = min(p1, p2), max(p1, p2)
                return mn, round((mn + mx) / 2, 2), mx
            except:
                pass
    # Angka tunggal
    try:
        v = float(raw)
        return (None, v, None) if v == 0 else (v, v, v)
    except:
        return None, None, None

# =====================
# 2. Koneksi ke DB
# =====================
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(
    f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;'
    f'UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes'
)
cr = conn.cursor()

# =====================
# 3. Proses & Update
# =====================
updated = 0
not_found = []

for _, row in df.iterrows():
    part = str(row[part_col]).strip()
    raw  = row[shrink_col]
    
    if not part or part.lower() in ('nan', ''):
        continue
    
    mn, std, mx = parse_shrink(raw)
    
    # Cari semua baris di SpsItemLists yang cocok dengan ItemList
    cr.execute(
        "SELECT Id FROM SpsItemLists WHERE ItemList = ?", part
    )
    ids = [r[0] for r in cr.fetchall()]
    
    if not ids:
        not_found.append(part)
        continue
    
    for item_id in ids:
        cr.execute("""
            UPDATE SpsItemLists
            SET LengthShrinkMin = ?,
                LengthShrinkStd = ?,
                LengthShrinkMax = ?
            WHERE Id = ?
        """, mn, std, mx, item_id)
        updated += 1

conn.commit()
conn.close()

# =====================
# 4. Laporan
# =====================
print(f"\nSelesai!")
print(f"Baris diupdate : {updated}")
print(f"Part tidak ada : {len(set(not_found))}")
if not_found:
    print("Contoh yg tidak ditemukan:", list(set(not_found))[:10])
