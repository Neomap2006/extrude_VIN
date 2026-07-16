import pyodbc
import sqlite3
import re
import decimal
import datetime

# 1. Connect to SQL Server
driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn_sql = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur_sql = conn_sql.cursor()

# 2. Connect to local SQLite
sqlite_db = "local_backup.sqlite"
conn_sqlite = sqlite3.connect(sqlite_db)
cur_sqlite = conn_sqlite.cursor()

# Get all tables
cur_sql.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG='prd_extrude_hose'")
tables = [row[0] for row in cur_sql.fetchall() if not row[0].startswith('__')]

print(f"Mulai copy {len(tables)} tabel ke SQLite ({sqlite_db})...")

for table in tables:
    # Get columns
    cur_sql.execute(f"SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{table}'")
    columns = cur_sql.fetchall()
    
    col_defs = []
    col_names = []
    for c in columns:
        col_name = f'"{c[0]}"'
        col_type = c[1].lower()
        col_names.append(col_name)
        
        # Mapping SQL Server types to SQLite types
        if 'int' in col_type or 'bit' in col_type:
            sqlite_type = 'INTEGER'
        elif 'decimal' in col_type or 'numeric' in col_type or 'float' in col_type or 'real' in col_type:
            sqlite_type = 'REAL'
        else:
            sqlite_type = 'TEXT'
            
        col_defs.append(f"{col_name} {sqlite_type}")
        
    # Create table in SQLite
    cur_sqlite.execute(f"DROP TABLE IF EXISTS \"{table}\"")
    create_stmt = f"CREATE TABLE \"{table}\" ({', '.join(col_defs)})"
    cur_sqlite.execute(create_stmt)
    
    # Get data from SQL Server
    try:
        cur_sql.execute(f"SELECT * FROM \"{table}\"")
        rows = cur_sql.fetchall()
        
        if not rows:
            print(f" - {table}: 0 baris (kosong)")
            continue
            
        # Insert into SQLite
        placeholders = ', '.join(['?' for _ in col_names])
        insert_stmt = f"INSERT INTO \"{table}\" ({', '.join(col_names)}) VALUES ({placeholders})"
        
        # Convert special types (like decimal or datetime) to string/float for SQLite
        cleaned_rows = []
        for row in rows:
            cleaned_row = []
            for val in row:
                if isinstance(val, decimal.Decimal):
                    cleaned_row.append(float(val))
                elif isinstance(val, (datetime.datetime, datetime.date, datetime.time)):
                    cleaned_row.append(str(val))
                else:
                    cleaned_row.append(val)
            cleaned_rows.append(cleaned_row)
            
        cur_sqlite.executemany(insert_stmt, cleaned_rows)
        print(f" - {table}: {len(cleaned_rows)} baris berhasil dicopy")
        
    except Exception as e:
        print(f" - {table}: Gagal! {e}")

conn_sqlite.commit()
conn_sqlite.close()
conn_sql.close()

print("\nBerhasil! Semua data sudah di-copy ke:", sqlite_db)
