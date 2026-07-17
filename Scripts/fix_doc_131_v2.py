import openpyxl
import pyodbc
import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

SQL_SERVER   = "10.14.149.34,1433"
SQL_DB       = "prd_extrude_hose"
SQL_USER     = "usrvelasto"
SQL_PASSWORD = "H1s@na2025!!"
EXCEL_FILE   = r"folder excel\Masterlist SPS CHS 2 Layer DIG.xlsx"
SHEET_NAME   = "Parameter Setting"

MAPPING_CHS2 = {
    'No': 2, 'DocumentNumber': 3, 'RevisionNumber': 6, 'Customer': 7, 'RevisionDate': 8,
    'Formulasi': 9, 'HoseType': 10, 'Dimensi': 11, 'Material': 12, 'InnerTube': 13,
    'OuterCover': 14, 'UseLimitsInner': 15, 'UseLimitsOuter': 16, 'Yarn': 17, 'PitchYarn': 18,
    'TensionYarnInner': 19, 'TensionYarnOuter': 20, 'Nipple': 21, 'TubeDie': 22, 'MiddleDie': 23,
    'CoverDie': 24, 'SpacerDie': 25, 'ADistance': 26, 'MeshDim1': 27, 'MeshScreen1': 28,
    'MeshDim2': 29, 'MeshScreen2': 30, 'HeadTemp1': 31, 'Cylinder1_1': 32, 'Cylinder2_1': 33,
    'Cylinder3_1': 34, 'ScrewTemp1': 35, 'HeadTemp2': 36, 'Cylinder1_2': 37, 'Cylinder2_2': 38,
    'Cylinder3_2': 39, 'ScrewTemp2': 40, 'ScrewSpeed1': 41, 'ScrewSpeed2': 42, 'FeedRollRatio1': 43,
    'FeedRollRatio2': 44, 'Pressure1': 45, 'Pressure2': 46, 'CurrentValue': 47, 'AmMeter': 48,
    'AmMeter2': 49, 'PresetValue': 50, 'ControlValue': 51, 'SpiralPitchSetting': 52,
    'SpiralPitchDisplay': 53, 'SpiralSpeed': 54, 'HoseSpeed': 55, 'UnsmoothSurface': 56,
    'MarkingSort': 57, 'TextMarkingMaterial': 58, 'MarkingColour': 59, 'ChillerWaterTemp': 60,
    'DancerPosition': 61, 'CaterpillarGap': 62, 'TakeUpConveyorSpeed': 63, 'CoolConveyorSpeed': 64,
    'ConveyorRatio': 65, 'ToleranceInner': 66, 'ToleranceOuter': 67, 'TebalInner': 68,
    'TebalTotal': 69, 'SelisihTebal': 70, 'InnerTarget': 71, 'InnerTol': 72, 'InnerLCL': 73,
    'InnerMin': 74, 'InnerUCL': 75, 'InnerMax': 76, 'ThickTarget': 77, 'ThickTol': 78,
    'ThickLCL': 79, 'ThickMin': 80, 'ThickUCL': 81, 'ThickMax': 82, 'TotalTarget': 83,
    'TotalTol': 84, 'TotalLCL': 85, 'TotalMin': 86, 'TotalUCL': 87, 'TotalMax': 88,
}

# The previous script mapping assumed combined fix file, where DocumentNumber was col 4.
# In Masterlist SPS CHS 2 Layer DIG.xlsx, based on the output of pandas:
# Unnamed: 3 = VI-SOP-PROD-131, so col 4 is No Document.
# Wait, let's just find the column by iterating header.
def get_column_mapping(ws, header_row=5):
    # This is a bit risky, let's just use the known mapping but shift if needed.
    pass

def v(ws, row, col):
    val = ws.cell(row=row, column=col).value
    s = str(val).strip() if val is not None else None
    return s if s and s != '-' else None

wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
ws = wb[SHEET_NAME]

# In the combined file, doc_val was col 4. Let's check col 4 here.
# Actually, let's iterate to find the actual column index of 'VI-SOP-PROD-131'.
doc_col = None
for c in range(1, 20):
    for r in range(1, 15):
        if v(ws, r, c) == 'No. Document':
            doc_col = c
            break

if not doc_col:
    # default to 4
    doc_col = 4

# The user screenshot showed col 7 is 'No. Rev.' and 9 is 'Revision Date' if doc_col is 4.
# Let's map exactly based on the previous working mapping for CHS 2!
MAPPING_CHS2 = {
    'No': 3, 'DocumentNumber': 4, 'RevisionNumber': 7, 'Customer': 8, 'RevisionDate': 9,
    'Formulasi': 10, 'HoseType': 11, 'Dimensi': 12, 'Material': 13, 'InnerTube': 14,
    'OuterCover': 15, 'UseLimitsInner': 16, 'UseLimitsOuter': 17, 'Yarn': 18, 'PitchYarn': 19,
    'TensionYarnInner': 20, 'TensionYarnOuter': 21, 'Nipple': 22, 'TubeDie': 23, 'MiddleDie': 24,
    'CoverDie': 25, 'SpacerDie': 26, 'ADistance': 27, 'MeshDim1': 28, 'MeshScreen1': 29,
    'MeshDim2': 30, 'MeshScreen2': 31, 'HeadTemp1': 32, 'Cylinder1_1': 33, 'Cylinder2_1': 34,
    'Cylinder3_1': 35, 'ScrewTemp1': 36, 'HeadTemp2': 37, 'Cylinder1_2': 38, 'Cylinder2_2': 39,
    'Cylinder3_2': 40, 'ScrewTemp2': 41, 'ScrewSpeed1': 42, 'ScrewSpeed2': 43, 'FeedRollRatio1': 44,
    'FeedRollRatio2': 45, 'Pressure1': 46, 'Pressure2': 47, 'CurrentValue': 48, 'AmMeter': 49,
    'AmMeter2': 50, 'PresetValue': 51, 'ControlValue': 52, 'SpiralPitchSetting': 53,
    'SpiralPitchDisplay': 54, 'SpiralSpeed': 55, 'HoseSpeed': 56, 'UnsmoothSurface': 57,
    'MarkingSort': 58, 'TextMarkingMaterial': 59, 'MarkingColour': 60, 'ChillerWaterTemp': 61,
    'DancerPosition': 62, 'CaterpillarGap': 63, 'TakeUpConveyorSpeed': 64, 'CoolConveyorSpeed': 65,
    'ConveyorRatio': 66, 'ToleranceInner': 67, 'ToleranceOuter': 68, 'TebalInner': 69,
    'TebalTotal': 70, 'SelisihTebal': 71, 'InnerTarget': 72, 'InnerTol': 73, 'InnerLCL': 74,
    'InnerMin': 75, 'InnerUCL': 76, 'InnerMax': 77, 'ThickTarget': 78, 'ThickTol': 79,
    'ThickLCL': 80, 'ThickMin': 81, 'ThickUCL': 82, 'ThickMax': 83, 'TotalTarget': 84,
    'TotalTol': 85, 'TotalLCL': 86, 'TotalMin': 87, 'TotalUCL': 88, 'TotalMax': 89,
}

latest_row_data = None
latest_row_idx = None
highest_rev = -1

for r in range(6, ws.max_row + 1):
    doc_val = v(ws, r, doc_col)
    if doc_val == 'VI-SOP-PROD-131':
        rev_val = v(ws, r, 7)
        try:
            rev_num = int(rev_val) if rev_val else 0
        except:
            rev_num = 0
            
        if rev_num > highest_rev:
            highest_rev = rev_num
            latest_row_idx = r
            
            row_data = {'IsActive': '1', 'Machine': 'CHS 2 Layer', 'MachineCode': 'CHS 2 Layer'}
            for field, col_idx in MAPPING_CHS2.items():
                row_data[field] = v(ws, r, col_idx)
            latest_row_data = row_data

wb.close()

if not latest_row_data:
    print("Document VI-SOP-PROD-131 tidak ditemukan di Excel.")
    sys.exit(1)

print(f"Mengambil data dari Excel baris ke-{latest_row_idx} (Revision: {latest_row_data['RevisionNumber']})")
print(f"Tolerance Inner dari Excel: {latest_row_data['ToleranceInner']}")

# Update DB
driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn_str = f"DRIVER={{{driver}}};SERVER={SQL_SERVER};DATABASE={SQL_DB};UID={SQL_USER};PWD={SQL_PASSWORD};TrustServerCertificate=yes;Encrypt=no;"
conn = pyodbc.connect(conn_str)
cur = conn.cursor()

cur.execute("DELETE FROM SpsNoDocs WHERE DocumentNumber = 'VI-SOP-PROD-131'")

fields = list(latest_row_data.keys())
placeholders = ', '.join(['?' for _ in fields])
col_names = ', '.join([f'[{f}]' for f in fields])
values = [latest_row_data[f] for f in fields]

sql = f'INSERT INTO SpsNoDocs ({col_names}) VALUES ({placeholders})'
cur.execute(sql, values)
conn.commit()
print("Record baru (Revisi terbaru) berhasil diinsert ke database!")

def parse_pm_value(s):
    if not s: return None, None, None
    s = s.replace(',', '.').strip()
    m = re.search(r'([0-9.]+)\s*[±+]\s*([0-9.]+)', s)
    if m:
        base = float(m.group(1))
        tol = float(m.group(2))
        return round(base - tol, 4), round(base, 4), round(base + tol, 4)
    return None, None, None

columns_to_parse = [
    'ToleranceInner', 'ToleranceOuter', 'TebalInner', 'TebalOuter', 'TebalTotal', 
    'TebalInnerMiddle', 'Nipple', 'TubeDie', 'CoverDie', 'MiddleDie', 'SpacerDie', 
    'ADistance', 'MeshDim1', 'MeshDim2', 'MeshDim3', 'HeadTemp1', 'HeadTemp2', 'HeadTemp3',
    'Cylinder1_1', 'Cylinder1_2', 'Cylinder1_3', 'Cylinder2_1', 'Cylinder2_2', 'Cylinder2_3',
    'Cylinder3_1', 'Cylinder3_2', 'Cylinder3_3', 'ScrewTemp1', 'ScrewTemp2', 'ScrewTemp3',
    'ScrewSpeed1', 'ScrewSpeed2', 'ScrewSpeed3', 'Pressure1', 'Pressure2', 'Pressure3',
    'HoseSpeed', 'TakeUpConveyorSpeed', 'ChillerWaterTemp', 'Feed1', 'Feed2', 'Feed3',
    'FeedRollRatio1', 'FeedRollRatio2', 'FeedRollRatio3', 'CaterpillarGap', 'SpiralSpeed',
    'AmMeter', 'AmMeter2', 'AmMeter3', 'PresetValue', 'ControlValue', 'SpiralPitchSetting',
    'SpiralPitchDisplay', 'CoolConveyorSpeed', 'CoolConveyorSpeed2', 'ConveyorRatio',
    'SelisihTebal', 'PitchYarn', 'DancerPosition', 'OdSensor', 'CuttingSpeed', 'ToleranceSpiralPitch'
]

cur.execute("SELECT * FROM SpsNoDocs WHERE DocumentNumber = 'VI-SOP-PROD-131'")
row = cur.fetchone()
cols = [c[0] for c in cur.description]
row_dict = dict(zip(cols, row))

updates = []
for col in columns_to_parse:
    if col in row_dict and row_dict[col]:
        mn, asli, mx = parse_pm_value(row_dict[col])
        if mn is not None:
            updates.append(f"[{col}_Min] = {mn}")
            updates.append(f"[{col}_Asli] = {asli}")
            updates.append(f"[{col}_Max] = {mx}")

if updates:
    update_sql = f"UPDATE SpsNoDocs SET {', '.join(updates)} WHERE DocumentNumber = 'VI-SOP-PROD-131'"
    cur.execute(update_sql)
    conn.commit()
    print("Nilai parsed (Min/STN/Max) berhasil diupdate!")
    
cur.execute("SELECT ToleranceInner, ToleranceInner_Min, ToleranceInner_Asli, ToleranceInner_Max FROM SpsNoDocs WHERE DocumentNumber = 'VI-SOP-PROD-131'")
res = cur.fetchone()
print(f"Verifikasi Database -> ToleranceInner: {res[0]}, Min: {res[1]}, STN: {res[2]}, Max: {res[3]}")
conn.close()
