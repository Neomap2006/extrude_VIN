import pyodbc
from datetime import datetime

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;', autocommit=True)
cur = conn.cursor()

# Get default backup directory
try:
    cur.execute("""
        EXEC master.dbo.xp_instance_regread 
            N'HKEY_LOCAL_MACHINE', 
            N'Software\Microsoft\MSSQLServer\MSSQLServer',N'BackupDirectory'
    """)
    row = cur.fetchone()
    backup_dir = row[1] if row else r"C:\Backup"
except Exception as e:
    backup_dir = r"C:\Backup"

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
backup_file = f"{backup_dir}\\prd_extrude_hose_{timestamp}.bak"

print(f"Mencoba melakukan backup ke server di path: {backup_file}")

try:
    cur.execute(f"BACKUP DATABASE prd_extrude_hose TO DISK = '{backup_file}' WITH FORMAT, MEDIANAME = 'SQLServerBackups', NAME = 'Full Backup of prd_extrude_hose';")
    print(f"Backup DATABASE berhasil! File tersimpan di server: {backup_file}")
except Exception as e:
    print(f"Backup SQL Server gagal (mungkin tidak ada akses): {e}")
    print("Akan membuat backup lokal (CSV) untuk tabel SpsNoDocs dan SpsItemLists...")

    import csv
    
    # Backup SpsNoDocs
    cur.execute("SELECT * FROM SpsNoDocs")
    docs = cur.fetchall()
    columns = [column[0] for column in cur.description]
    with open('backup_SpsNoDocs.csv', 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        for row in docs:
            writer.writerow(row)
    print(f"Lokal backup berhasil: backup_SpsNoDocs.csv ({len(docs)} baris)")
    
    # Backup SpsItemLists
    cur.execute("SELECT * FROM SpsItemLists")
    items = cur.fetchall()
    columns = [column[0] for column in cur.description]
    with open('backup_SpsItemLists.csv', 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        for row in items:
            writer.writerow(row)
    print(f"Lokal backup berhasil: backup_SpsItemLists.csv ({len(items)} baris)")
