import pandas as pd
import re

file_path = r"d:\Project\extrude\folder excel\EXPORT_PIVOT.xlsx"
df = pd.read_excel(file_path, header=None)

# Find header row
header_row_idx = 0
for i in range(10):
    row_vals = df.iloc[i].astype(str).str.upper().tolist()
    if any('VIN' in str(v) for v in row_vals) or any('SHRINK' in str(v) for v in row_vals):
        header_row_idx = i
        break

print(f"Header found at row {header_row_idx}")
df.columns = df.iloc[header_row_idx]
df = df.iloc[header_row_idx+1:].reset_index(drop=True)

# Let's find the Part No VIN column and Length shrink (mm) column
part_col = [c for c in df.columns if 'PART NO' in str(c).upper() or 'VIN' in str(c).upper()]
shrink_col = [c for c in df.columns if 'SHRINK' in str(c).upper()]

if part_col and shrink_col:
    print(f"\nUsing columns: {part_col[0]} and {shrink_col[0]}")
    sample_df = df[[part_col[0], shrink_col[0]]].dropna().head(10)
    
    print("\n--- SAMPLE DATA ---")
    for _, row in sample_df.iterrows():
        part = str(row[part_col[0]]).strip()
        shrink = str(row[shrink_col[0]]).strip()
        
        # Parse logic
        shrink_min = None
        shrink_std = None
        shrink_max = None
        
        # Common formats:
        # "501±2" -> Std=501, Min=499, Max=503
        # "187±1" -> Std=187, Min=186, Max=188
        # "130-135" -> Min=130, Max=135, Std=132.5
        # "0" -> Std=0
        
        if '±' in shrink:
            parts = shrink.split('±')
            if len(parts) == 2:
                try:
                    std = float(parts[0].replace(',','.'))
                    tol = float(parts[1].replace(',','.'))
                    shrink_std = std
                    shrink_min = std - tol
                    shrink_max = std + tol
                except: pass
        elif '-' in shrink and not shrink.startswith('-'):
            parts = shrink.split('-')
            if len(parts) == 2:
                try:
                    p1 = float(parts[0].replace(',','.'))
                    p2 = float(parts[1].replace(',','.'))
                    shrink_min = min(p1, p2)
                    shrink_max = max(p1, p2)
                    shrink_std = (shrink_min + shrink_max) / 2
                except: pass
        else:
            try:
                v = float(shrink.replace(',','.'))
                shrink_std = v
            except: pass
            
        print(f"Part: {part:10} | Raw: {shrink:15} => Min: {shrink_min}, Std: {shrink_std}, Max: {shrink_max}")
else:
    print("Could not find the expected columns.")
