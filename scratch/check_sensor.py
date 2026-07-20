import pyodbc
from datetime import datetime

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn_str = f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;'
conn = pyodbc.connect(conn_str)
cur = conn.cursor()

print("--- SENSOR DATA CHECK ---")
# Check total records
cur.execute("SELECT COUNT(*) FROM SensorIngestLogs")
print("Total SensorIngestLogs count:", cur.fetchone()[0])

# Check today's records based on IngestedAt
cur.execute("""
    SELECT COUNT(*) 
    FROM SensorIngestLogs 
    WHERE CAST(IngestedAt AS DATE) = CAST(GETDATE() AS DATE)
""")
print("Today's IngestedAt records count:", cur.fetchone()[0])

# Check today's records based on SensorTimestamp
cur.execute("""
    SELECT COUNT(*) 
    FROM SensorIngestLogs 
    WHERE CAST(SensorTimestamp AS DATE) = CAST(GETDATE() AS DATE)
""")
print("Today's SensorTimestamp records count:", cur.fetchone()[0])

# Check latest 5 records
print("\nLatest 5 Records:")
cur.execute("""
    SELECT TOP 5 Id, DeviceId, MachineCode, MetricType, MetricValue, Quality, SensorTimestamp, IngestedAt 
    FROM SensorIngestLogs 
    ORDER BY Id DESC
""")
for r in cur.fetchall():
    print(r)
