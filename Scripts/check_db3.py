import pyodbc
driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn_str = f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;'
conn = pyodbc.connect(conn_str)
cur = conn.cursor()
cur.execute("SELECT * FROM SpsNoDocs WHERE DocumentNumber = 'SOP/PROD/HOSE/SPS/3/1'")
columns = [column[0] for column in cur.description]
row = cur.fetchone()
if row:
    for col, val in zip(columns, row):
        if val is not None:
            print(f'{col}: {val}')
else:
    print('Not found')
