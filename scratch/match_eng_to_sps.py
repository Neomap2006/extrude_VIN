import openpyxl
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r'd:\Project\extrude\folder excel\Masterlist SPS Double Layer.xlsx', data_only=True, read_only=True)
ws = wb['Parameter Setting']

# Col index: 2=Machine, 3=DocumentNumber, 48=EPR doc, 50=ITEM
eng_rows = {}   # item -> {eng_doc, epr_doc}
sps_rows = {}   # item -> {sps_doc, machine}

for row in ws.iter_rows(values_only=True):
    row_list = list(row)
    doc = str(row_list[3] if len(row_list) > 3 else '').strip()
    item_raw = str(row_list[50] if len(row_list) > 50 else '').strip()
    epr_doc = str(row_list[48] if len(row_list) > 48 else '').strip()
    machine_col = str(row_list[2] if len(row_list) > 2 else '').strip()
    
    if not item_raw or item_raw in ['nan', '-', 'ITEM', 'None']:
        continue
    
    # Split by comma only (to keep "NA 204" as one item)
    items = [x.strip() for x in item_raw.split(',') if x.strip()]
    
    if 'SOP/ENG' in doc:
        for itm in items:
            eng_rows[itm] = {'eng_doc': doc, 'epr_doc': epr_doc}
    elif 'SOP/PROD' in doc:
        for itm in items:
            sps_rows[itm] = {'sps_doc': doc, 'machine': machine_col}

wb.close()

print(f"Items dari SOP/ENG: {len(eng_rows)}")
print(f"Items dari SPS rows: {len(sps_rows)}")
print()
print("=== SPS items tersedia ===")
for itm, v in sorted(sps_rows.items()):
    print(f"  {itm:<20} -> {v['sps_doc']:<35} | {v['machine']}")

print()
print("=== HASIL PENCARIAN MESIN UNTUK SOP/ENG ===")
print(f"{'ENG Document':<25} | {'ITEM':<20} | {'SPS Document':<35} | {'Machine'}")
print("-" * 110)

eng_doc_results = {}
for item, eng_info in sorted(eng_rows.items()):
    if item in sps_rows:
        sps_info = sps_rows[item]
        print(f"{eng_info['eng_doc']:<25} | {item:<20} | {sps_info['sps_doc']:<35} | {sps_info['machine']}")
        eng_doc_results[eng_info['eng_doc']] = sps_info['machine']
    else:
        print(f"{eng_info['eng_doc']:<25} | {item:<20} | (TIDAK DITEMUKAN)                   |")

print()
print("=== KESIMPULAN: ENG Doc -> Machine ===")
for doc, mc in sorted(eng_doc_results.items()):
    print(f"  {doc} -> {mc}")
