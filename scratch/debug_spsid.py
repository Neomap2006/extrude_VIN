import pyodbc
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

# Cek SpsId distribution
cr.execute("SELECT SpsId, COUNT(*) as cnt FROM ProductionReports GROUP BY SpsId ORDER BY SpsId")
rows = cr.fetchall()
print("SpsId distribution:", [(r[0], r[1]) for r in rows[:20]])

# Cek apakah ada DocumentNumber matching
cr.execute("SELECT TOP 5 Id, DocumentNumber, SpsId, ItemCode, ProductionDate FROM ProductionReports WHERE DocumentNumber LIKE '%SPS%'")
rows2 = cr.fetchall()
print("\nSample DocumentNumber in ProductionReports:")
for r in rows2:
    print(r)
