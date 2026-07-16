import pyodbc
import sys
sys.stdout.reconfigure(encoding='utf-8')

driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
conn = pyodbc.connect(f'DRIVER={{{driver}}};SERVER=10.14.149.34,1433;DATABASE=prd_extrude_hose;UID=usrvelasto;PWD=H1s@na2025!!;TrustServerCertificate=yes;Encrypt=no;')
cur = conn.cursor()

# Fix: SPS/24/* dan SPS/1/* yang masih kosong -> CHS1
cur.execute("""
    UPDATE SpsNoDocs SET MachineCode = 'CHS1'
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%'
    AND (MachineCode IS NULL OR MachineCode = '')
""")
print(f"Updated {cur.rowcount} records -> CHS1")
conn.commit()

# Verify final distribution
cur.execute("""
    SELECT MachineCode, COUNT(*) as cnt
    FROM SpsNoDocs
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%'
    GROUP BY MachineCode
    ORDER BY MachineCode
""")
print("\nFinal MachineCode distribution:")
for row in cur.fetchall():
    print(f"  [{row[0]}] -> {row[1]} dokumen")

# ============================================================
# Regenerate report per mesin
# ============================================================
cur.execute("""
    SELECT DocumentNumber, MachineCode, Customer, HoseType, Material, InnerTube, OuterCover, MiddleTube
    FROM SpsNoDocs
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%'
    ORDER BY MachineCode, DocumentNumber
""")
docs = cur.fetchall()

cur.execute("""
    SELECT DocumentNumber, ItemList
    FROM SpsItemLists
    WHERE DocumentNumber LIKE 'SOP/PROD/HOSE/SPS/%'
    ORDER BY DocumentNumber, ItemList
""")
items_map = {}
for doc_num, item in cur.fetchall():
    items_map.setdefault(doc_num, []).append(item)

artifact_path = r"C:\Users\megi.vlt00553\.gemini\antigravity-ide\brain\a52d2d79-d482-4912-af40-3b1539b302ac\data_verification_report.md"

machines = {}
for row in docs:
    mc = row[1] or 'UNKNOWN'
    machines.setdefault(mc, []).append(row)

with open(artifact_path, "w", encoding="utf-8") as f:
    f.write("# Verifikasi Data Master SPS — Per Mesin\n\n")
    f.write("> [!NOTE]\n")
    f.write("> Laporan ini menampilkan semua dokumen SPS baru yang telah diimport, dikelompokkan per mesin. Gunakan untuk cross-check dengan file Excel `Masterlist_Digitalisasi_Combined_Fix.xlsx`.\n\n")
    f.write(f"**Total Dokumen:** {len(docs)} | **Total Mesin:** {len(machines)}\n\n")

    for mc in sorted(machines.keys()):
        mc_docs = machines[mc]
        total_items = sum(len(items_map.get(d[0], [])) for d in mc_docs)
        f.write(f"## Mesin: {mc} — {len(mc_docs)} Dokumen, {total_items} Item Lists\n\n")
        f.write("| No | No Document | Customer | Hose Type | Material | Inner | Outer | Middle | Item Lists |\n")
        f.write("|---|---|---|---|---|---|---|---|---|\n")
        for i, d in enumerate(mc_docs, 1):
            doc_num  = d[0]
            customer = d[2] or '-'
            hose     = d[3] or '-'
            mat      = d[4] or '-'
            inner    = d[5] or '-'
            outer    = d[6] or '-'
            middle   = d[7] or '-'
            items    = ", ".join(items_map.get(doc_num, []))
            if not items:
                items = "*(tidak ada)*"
            f.write(f"| {i} | `{doc_num}` | {customer} | {hose} | {mat} | {inner} | {outer} | {middle} | {items} |\n")
        f.write("\n---\n\n")

print("\nReport generated!")
