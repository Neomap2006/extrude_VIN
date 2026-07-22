import pyodbc
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()
try:
    cr.execute("SELECT Machine, Nipple, TubeDie, CoverDie, HoseType FROM SpsMasters WHERE Machine = 'Non-CHS (DL)' AND Nipple IS NOT NULL AND Nipple != ''")
    rows = cr.fetchall()
    print(f"Found {len(rows)} records with Nipple for Non-CHS (DL):")
    for row in rows[:5]:
        print(row)
except Exception as e:
    print("Error checking SpsMasters:", e)
