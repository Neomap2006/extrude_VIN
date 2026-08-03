import pandas as pd
file_path = r"d:\Project\extrude\folder excel\EXPORT_PIVOT.xlsx"
df = pd.read_excel(file_path, sheet_name="Sum", header=None)
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
shrink_col = [c for c in df.columns if 'SHRINK' in c.upper()][0]

row = df[df[part_col].astype(str).str.strip().str.upper() == 'PM1110']
if not row.empty:
    print("Shrink from excel for PM1110:")
    print(repr(row[shrink_col].iloc[0]))
