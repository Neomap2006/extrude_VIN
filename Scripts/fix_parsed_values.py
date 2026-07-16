import pyodbc
import re

def parse_plus_minus(input_str):
    if not input_str:
        return None, None, None
    s = str(input_str).strip().replace(" ", "").replace(",", ".")
    
    if s.lower().startswith('max') or s.lower().startswith('maks'):
        m = re.search(r'\d+(\.\d+)?', s)
        if m: return None, None, float(m.group())
            
    if s.lower().startswith('min'):
        m = re.search(r'\d+(\.\d+)?', s)
        if m: return float(m.group()), None, None
            
    for sym in ['\u00b1', '±', '\ufffd', '+/-', '+-', '~']:
        if sym in s:
            parts = s.split(sym)
            if len(parts) == 2:
                try:
                    std = float(parts[0])
                    tol = float(parts[1])
                    return std - tol, std, std + tol
                except:
                    pass
                    
    # Format X-Y: range min-max (e.g. "2-5", "13.5-16") → Min=X, Max=Y, Asli=None
    # Must match "digit-digit" pattern, avoiding negative numbers (which start with -)
    m_range = re.match(r'^(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)$', s)
    if m_range:
        try:
            lo = float(m_range.group(1))
            hi = float(m_range.group(2))
            return lo, None, hi
        except:
            pass

    try:
        val = float(s)
        return None, val, None
    except:
        pass
        
    return None, None, None

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

# Get all base columns that have _Min, _Asli, _Max
cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'SpsNoDocs'")
all_cols = [r[0] for r in cur.fetchall()]
base_cols = [c[:-5] for c in all_cols if c.endswith('_Asli')]

# We only want to update the records we just inserted (SOP/PROD/HOSE/SPS/1, 2, 3)
cur.execute("SELECT DocumentNumber, " + ", ".join(base_cols) + " FROM SpsNoDocs WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%'")
records = cur.fetchall()

print(f"Found {len(records)} records to process.")
updates = 0

for row in records:
    doc_num = row[0]
    update_parts = []
    update_params = []
    
    for i, base_col in enumerate(base_cols):
        val = row[i+1] # +1 because doc_num is at index 0
        if val is not None and str(val).strip() != '':
            cmin, casli, cmax = parse_plus_minus(val)
            if cmin is not None or casli is not None or cmax is not None:
                if cmin is not None:
                    update_parts.append(f"{base_col}_Min = ?")
                    update_params.append(cmin)
                if casli is not None:
                    update_parts.append(f"{base_col}_Asli = ?")
                    update_params.append(casli)
                if cmax is not None:
                    update_parts.append(f"{base_col}_Max = ?")
                    update_params.append(cmax)
                    
    if update_parts:
        sql = f"UPDATE SpsNoDocs SET {', '.join(update_parts)} WHERE DocumentNumber = ?"
        update_params.append(doc_num)
        cur.execute(sql, update_params)
        updates += 1
        
conn.commit()
print(f"Updated {updates} records successfully.")
