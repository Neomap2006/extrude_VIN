import openpyxl
import pyodbc
import glob
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
SQL_SERVER   = "10.14.149.34,1433"
SQL_DB       = "prd_extrude_hose"
SQL_USER     = "usrvelasto"
SQL_PASSWORD = "H1s@na2025!!"

print("Mengambil data dari Database...")
try:
    driver = [d for d in pyodbc.drivers() if 'SQL Server' in d][-1]
    conn_str = f"DRIVER={{{driver}}};SERVER={SQL_SERVER};DATABASE={SQL_DB};UID={SQL_USER};PWD={SQL_PASSWORD};TrustServerCertificate=yes;Encrypt=no;"
    conn = pyodbc.connect(conn_str)
    cur = conn.cursor()
    cur.execute("SELECT DocumentNumber, RevisionNumber FROM SpsNoDocs WHERE DocumentNumber IS NOT NULL")
    db_docs = {}
    for doc, rev in cur.fetchall():
        doc_str = str(doc).strip()
        if not doc_str: continue
        try:
            rev_num = int(rev) if rev else 0
        except ValueError:
            # If revision is a date, it might fail. E.g., '12-05-18' was inserted as revision.
            # Wait, our previous check showed RevisionNumber was '5' for the new one, and '12-05-18' for the old one?
            # Actually, in the old one it was probably incorrectly mapped. Let's just try to parse int.
            # If it fails, we set DB rev to 0 so it will be replaced by a valid integer revision anyway.
            rev_num = 0
        db_docs[doc_str] = rev_num
    conn.close()
    print(f"Berhasil memuat {len(db_docs)} dokumen dari Database.\\n")
except Exception as e:
    print(f"Error Database: {e}")
    sys.exit(1)

# Now scan Excel files
excel_files = [
    r"folder excel\Masterlist_Digitalisasi_Combined_Fix.xlsx",
    r"folder excel\Masterlist SPS CHS 2 Layer DIG.xlsx",
    r"folder excel\Masterlist SPS CHS 3 Layer DIG.xlsx",
    r"folder excel\Masterlist SPS Double Layer.xlsx",
    r"folder excel\Masterlist SPS Double Layer_Digitalisasi.xlsx"
]

excel_max_revs = {}

print("Memindai file-file Excel...")
for f in excel_files:
    if not os.path.exists(f):
        continue
    print(f"  Membaca: {os.path.basename(f)}")
    try:
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # Identify columns for DocumentNumber and RevisionNumber
            # We assume header is within first 10 rows.
            doc_col = -1
            rev_col = -1
            
            for r in range(1, 15):
                row_cells = ws[r]
                for idx, cell in enumerate(row_cells):
                    val = str(cell.value).strip().lower() if cell.value else ""
                    if val == "no. document":
                        doc_col = idx
                    elif val in ["no. rev.", "no. rev", "revision"]:
                        rev_col = idx
                if doc_col != -1 and rev_col != -1:
                    break
            
            # if we couldn't find header exactly, fallback to common columns:
            if doc_col == -1: doc_col = 3  # 0-indexed for col 4
            if rev_col == -1: rev_col = 6  # 0-indexed for col 7
            
            for row in ws.iter_rows(min_row=5, values_only=True):
                if doc_col < len(row) and row[doc_col]:
                    doc_val = str(row[doc_col]).strip()
                    if doc_val and doc_val.startswith("VI-SOP") or doc_val.startswith("SOP/"):
                        rev_val = row[rev_col] if rev_col < len(row) else None
                        try:
                            rev_num = int(rev_val) if rev_val else 0
                        except:
                            rev_num = 0
                            
                        if doc_val not in excel_max_revs or rev_num > excel_max_revs[doc_val]['rev']:
                            excel_max_revs[doc_val] = {
                                'rev': rev_num,
                                'file': os.path.basename(f),
                                'sheet': sheet_name
                            }
        wb.close()
    except Exception as e:
        print(f"Error membaca {f}: {e}")

print("\\n--- HASIL ANALISA ---")
outdated = []
for doc, db_rev in db_docs.items():
    if doc in excel_max_revs:
        ex_rev = excel_max_revs[doc]['rev']
        if ex_rev > db_rev:
            outdated.append({
                'doc': doc,
                'db_rev': db_rev,
                'ex_rev': ex_rev,
                'file': excel_max_revs[doc]['file'],
                'sheet': excel_max_revs[doc]['sheet']
            })

if outdated:
    print(f"Ditemukan {len(outdated)} dokumen yang tertinggal revisi (database lebih lama dari Excel):\\n")
    for item in outdated:
        print(f"[{item['doc']}]")
        print(f"  - Database Rev: {item['db_rev']}")
        print(f"  - Excel Rev   : {item['ex_rev']} (ditemukan di {item['file']} sheet '{item['sheet']}')")
else:
    print("Semua dokumen di database sudah memakai revisi terbaru dari Excel!")
