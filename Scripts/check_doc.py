import pyodbc
driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()
cur.execute("SELECT DocumentNumber, Machine FROM SpsNoDocs WHERE DocumentNumber = 'VI-SOP-PROD-131'")
row = cur.fetchone()
if row:
    print(f'Ada! Document: {row[0]}, Machine: {row[1]}')
else:
    print('Tidak ditemukan di database.')
