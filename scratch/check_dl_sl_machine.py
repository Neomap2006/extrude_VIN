import pyodbc

d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

print("=== SEMUA TIPE MESIN DI DATABASE ===")
cr.execute("SELECT DISTINCT Machine, COUNT(*) as Jml FROM SpsNoDocs GROUP BY Machine ORDER BY Machine")
for r in cr.fetchall():
    print(f"  '{r[0]}' -> {r[1]} records")

print("\n=== DETAIL Non-CHS (DL) dan (SL) ===")
cr.execute("SELECT No, Machine, MachineCode, HoseType, DocumentNumber FROM SpsNoDocs WHERE Machine LIKE 'Non-CHS%' ORDER BY Machine")
rows = cr.fetchall()
print(f"Total: {len(rows)} rows\n")
print(f"{'No':<5} {'Machine':<20} {'MachineCode':<15} {'HoseType':<20} {'Doc':<35}")
print("-"*100)
for r in rows:
    print(f"{str(r[0]):<5} {str(r[1]):<20} {str(r[2] or '-'):<15} {str(r[3] or '-'):<20} {str(r[4] or '-'):<35}")
