import pyodbc
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()
cr.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'SpsNoDocs'")
columns = [r[0] for r in cr.fetchall()]
matches = [c for c in columns if 'PANJANG' in c.upper() or 'LENGTH' in c.upper() or 'STD' in c.upper() or 'LONG' in c.upper()]
print('Columns found:', matches)
print('All columns:', columns)
