import pyodbc

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

# Check Nipple values in the new records
cur.execute("""
SELECT DocumentNumber, Nipple, Nipple_Min, Nipple_Asli, Nipple_Max
FROM SpsNoDocs 
WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%' 
ORDER BY DocumentNumber
""")
print(f"{'DocumentNumber':<35} {'Nipple':<15} {'Min':<10} {'Asli':<10} {'Max':<10}")
print("-"*80)
for row in cur.fetchall():
    doc, nipple, mn, asli, mx = row
    print(f"{doc:<35} {str(nipple or ''):<15} {str(mn or ''):<10} {str(asli or ''):<10} {str(mx or ''):<10}")
