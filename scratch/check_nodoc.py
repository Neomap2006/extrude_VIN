import pyodbc
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()
try:
    cr.execute("SELECT TOP 1 * FROM SpsNoDocs")
    cols = [column[0] for column in cr.description]
    machine_cols = [c for c in cols if 'MACH' in c.upper()]
    nipple_cols = [c for c in cols if 'NIPPLE' in c.upper()]
    
    if machine_cols and nipple_cols:
        m_col = machine_cols[0]
        n_col = nipple_cols[0]
        
        doc_col = next((c for c in cols if 'DOC' in c.upper() or 'NUMBER' in c.upper() or 'SPS' in c.upper()), cols[0])
        
        # Check all DL1 machines without nipple
        query = f"SELECT {doc_col}, {m_col}, {n_col}, HoseType FROM SpsNoDocs WHERE {m_col} LIKE '%DL1%' OR {m_col} = 'Non-CHS (DL)'"
        cr.execute(query)
        rows = cr.fetchall()
        
        dl1_count = len(rows)
        no_nipple = [r for r in rows if r[2] is None or str(r[2]).strip() == '' or str(r[2]).strip() == '--' or str(r[2]).strip() == '0']
        
        print(f"Total DL1 records in SpsNoDocs: {dl1_count}")
        print(f"DL1 records WITHOUT Nipple: {len(no_nipple)}")
        for r in no_nipple[:10]:
            print(r)
except Exception as e:
    print("Error checking SpsNoDocs:", e)
