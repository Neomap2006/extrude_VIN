import pyodbc

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

def check_table(table_name):
    print(f"==== {table_name} ====")
    cr.execute(f"SELECT TOP 1 * FROM {table_name}")
    cols = [column[0] for column in cr.description]
    doc_cols = [c for c in cols if 'DOC' in c.upper() or 'NUMBER' in c.upper() or 'SPS' in c.upper() or 'NO' in c.upper()]
    
    mid_cols = [c for c in cols if 'MIDDLE' in c.upper() or 'INNERMID' in c.upper() or 'TARGET' in c.upper() or 'DIE' in c.upper()]
    
    for doc_col in doc_cols:
        query = f"SELECT {', '.join(mid_cols)} FROM {table_name} WHERE {doc_col} = 'SOP/PROD/HOSE/SPS/2/14'"
        try:
            cr.execute(query)
            rows = cr.fetchall()
            if rows:
                print(f"Found match using {doc_col}:")
                for row in rows:
                    for i, val in enumerate(row):
                        print(f"  {mid_cols[i]}: {val}")
                return
        except Exception:
            pass
    print("Not found")

check_table("SpsMasters")
check_table("SpsNoDocs")
