import pyodbc
import pandas as pd
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(
    f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;'
    f'UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes'
)
df = pd.read_sql("SELECT LengthShrinkMin, LengthShrinkStd, LengthShrinkMax FROM SpsItemLists WHERE ItemList='PM1110'", conn)
print(df)
