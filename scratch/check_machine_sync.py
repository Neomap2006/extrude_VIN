import pyodbc
import pandas as pd
import glob

# 1. Fetch from Database
print("=== DATABASE: Non-CHS (DL/SL) ===")
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()
cr.execute("""
    SELECT No, Machine, DocumentNumber 
    FROM SpsNoDocs 
    WHERE Machine LIKE 'Non-CHS%'
""")
db_rows = cr.fetchall()
db_docs = [r[2] for r in db_rows if r[2]]
print(f"Found {len(db_rows)} rows in SpsNoDocs with Machine like 'Non-CHS%'")
for r in db_rows[:5]:
    print(" ", r)

# 2. Search in Excel Files
print("\n=== EXCEL: Searching for specific machine names ===")
excel_files = glob.glob(r'd:\Project\extrude\folder excel\Masterlist*.xlsx')

found_docs = set()
for f in excel_files:
    if '~$' in f: continue
    try:
        df_dict = pd.read_excel(f, sheet_name=None, header=None) # read without header first
        for sheet_name, df in df_dict.items():
            # Search for the document numbers
            for doc in db_docs[:5]: # just check first 5 to see structure
                mask = df.astype(str).apply(lambda col: col.str.contains(str(doc), case=False, na=False))
                if mask.any().any():
                    print(f"[FOUND] {doc} in {f.split(chr(92))[-1]} -> Sheet: {sheet_name}")
                    
                    # Get the row
                    rows_with_doc = df[mask.any(axis=1)]
                    for idx, row in rows_with_doc.iterrows():
                        # Let's print non-null values in this row to see if there's a specific machine name
                        vals = [str(x) for x in row if pd.notna(x)]
                        print("   ROW DATA:", [v for v in vals if v.strip() != ''])
                    found_docs.add(doc)
    except Exception as e:
        print(f"Error reading {f}: {e}")

if not found_docs:
    print("Could not find any of the Non-CHS docs in the Excel files.")
