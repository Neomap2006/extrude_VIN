import pyodbc
d = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={d};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes')
cr = conn.cursor()

DOC = 'SOP/PROD/HOSE/SPS/2/9'

# 1. Standar dari SpsNoDocs
print(f"=== STANDAR SPS: {DOC} ===")
cr.execute("""
    SELECT No, SpiralSpeed_Min, SpiralSpeed_Asli, SpiralSpeed_Max,
           SpiralPitchSetting_Min, SpiralPitchSetting_Asli, SpiralPitchSetting_Max,
           HoseType, Dimensi, Machine
    FROM SpsNoDocs WHERE DocumentNumber = ?
""", DOC)
sps = cr.fetchone()
if not sps:
    print("TIDAK DITEMUKAN di SpsNoDocs!")
    exit()

sps_id = sps[0]
spd_min, spd_std, spd_max = sps[1], sps[2], sps[3]
pitch_min, pitch_std, pitch_max = sps[4], sps[5], sps[6]
hose_type, dimensi, machine = sps[7], sps[8], sps[9]

print(f"  SpsNo      : {sps_id}")
print(f"  HoseType   : {hose_type}")
print(f"  Dimensi    : {dimensi}")
print(f"  Machine    : {machine}")
print(f"  Spiral Speed: {spd_min} | {spd_std} | {spd_max}  (Min | STD | Max)")
print(f"  Pitch Set  : {pitch_min} | {pitch_std} | {pitch_max}  (Min | STD | Max)")

# 2. Cek kolom ProductionReports
cr.execute("SELECT TOP 1 * FROM ProductionReports")
rpt_cols = [c[0] for c in cr.description]
sps_cols = [c for c in rpt_cols if 'sps' in c.lower() or 'item' in c.lower() or 'doc' in c.lower()]
print(f"\n  Kolom SPS di ProductionReports: {sps_cols}")

# 3. Cari laporan berdasarkan ItemCode
print(f"\n=== RIWAYAT PRODUKSI (DocumentNumber: {DOC}) ===")
cr.execute("""
    SELECT pr.Id, pr.ProductionDate, pr.Shift, pr.CreatedBy, pr.Status,
           pr.InitSpiralSpeed, pr.InitSpiralPitchDisplay
    FROM ProductionReports pr
    WHERE pr.DocumentNumber = ?
    ORDER BY pr.ProductionDate DESC
""", DOC)
reports = cr.fetchall()
print(f"  Total laporan: {len(reports)}")

if reports:
    out_count = 0
    print(f"\n  {'Tgl':<12} {'Shift':<6} {'Operator':<22} {'InitSpeed':<12} {'PitchDisp':<12} Status")
    print(f"  {'-'*85}")
    for r in reports:
        pid, tgl, shift, op, status, sp_speed, sp_pitch = r
        spd_str = str(sp_speed) if sp_speed is not None else '-'
        pitch_str = str(sp_pitch) if sp_pitch is not None else '-'

        flag = ""
        if sp_speed is not None:
            try:
                v = float(str(sp_speed))
                if spd_max and v > float(spd_max):
                    flag = " [OVER MAX]"
                    out_count += 1
                elif spd_min and v < float(spd_min):
                    flag = " [UNDER MIN]"
                    out_count += 1
            except: pass

        tgl_str = str(tgl)[:10] if tgl else '-'
        op_str = str(op)[:20] if op else '-'
        st_str = str(status)[:10] if status else '-'
        print(f"  {tgl_str:<12} {str(shift or '-'):<6} {op_str:<22} {spd_str:<12} {pitch_str:<12} {st_str}{flag}")

    pct = 100*out_count//len(reports) if reports else 0
    print(f"\n  ⚠️  Keluar standar: {out_count}/{len(reports)} laporan ({pct}%)")

# 4. Sensor readings
print(f"\n=== SENSOR READINGS (SpiralSpeed) ===")
cr.execute("""
    SELECT rd.SpiralSpeed, COUNT(*) as cnt
    FROM ProductionReadings rd
    JOIN ProductionReports rpt ON rd.ProductionReportId = rpt.Id
    WHERE rpt.DocumentNumber = ?
      AND rd.SpiralSpeed IS NOT NULL
    GROUP BY rd.SpiralSpeed
    ORDER BY COUNT(*) DESC
""", DOC)
readings = cr.fetchall()
if readings:
    print(f"  Distribusi SpiralSpeed dari sensor:")
    for val, cnt in readings[:20]:
        flag = ""
        try:
            v = float(str(val))
            if spd_max and v > float(spd_max):   flag = "  [OVER]"
            elif spd_min and v < float(spd_min): flag = "  [UNDER]"
        except: pass
        bar = '█' * min(int(cnt/5), 30)
        print(f"    {str(val):>8} rpm  x{cnt:>6}  {bar}{flag}")
else:
    print("  Tidak ada data sensor")

conn.close()
