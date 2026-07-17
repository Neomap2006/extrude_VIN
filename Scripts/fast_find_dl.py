import openpyxl

f = r'folder excel\Masterlist SPS Double Layer.xlsx'
print(f"Membaca {f}...")
wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
ws = wb['Parameter Setting']

doc_to_find = 'SOP/PROD/HOSE/EPR/24/10/13'
found = False

for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
    for c_idx, val in enumerate(row, start=1):
        if str(val).strip() == doc_to_find:
            print(f"Found {doc_to_find} at row {r_idx}, col {c_idx}")
            found = True
            break
    if found:
        break

if not found:
    print(f"Could not find {doc_to_find}")
