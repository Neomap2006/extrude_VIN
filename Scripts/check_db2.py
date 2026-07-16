import pyodbc
driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn_str = f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;'
conn = pyodbc.connect(conn_str)
cur = conn.cursor()
cur.execute("SELECT COUNT(1) FROM SpsItemLists WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/1/1%'")
print('SpsItemLists matches:', cur.fetchone()[0])
cur.execute("SELECT TOP 5 ItemList, DocumentNumber FROM SpsItemLists WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/1/1%'")
for r in cur.fetchall():
    print(r)
