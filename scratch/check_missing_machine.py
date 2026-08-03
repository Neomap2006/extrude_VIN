import pandas as pd

# Part yang dicari
missing_parts = ['HN0490', 'HN0590', 'HN0540', 'HN0530', 'PM1110', 'HN0300', 'HN0550', 'HN0250', 'HN0560']

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

# Cari kolom part dan extrude
part_col = [c for c in df.columns if 'VIN' in c.upper() or 'PART NO' in c.upper()][0]
extrude_col = [c for c in df.columns if 'EXTRUDE' in c.upper()]

if extrude_col:
    extrude_col = extrude_col[0]
    print(f"Menggunakan kolom part: '{part_col}', kolom mesin: '{extrude_col}'\n")
    print(f"{'Part No VIN':<15} | {'Mesin (Extrude)':<20}")
    print("-" * 40)
    
    # Filter data yang dicari
    for part in missing_parts:
        # Cari baris yang cocok (case-insensitive & handle whitespace)
        row = df[df[part_col].astype(str).str.strip().str.upper() == part.upper()]
        if not row.empty:
            mesin = str(row[extrude_col].iloc[0]).strip()
            print(f"{part:<15} | {mesin:<20}")
        else:
            print(f"{part:<15} | TIDAK DITEMUKAN DI EXCEL")
else:
    print("Kolom dengan nama 'EXTRUDE' tidak ditemukan. Berikut daftar kolom yang ada:")
    for c in df.columns:
        print("-", c)
