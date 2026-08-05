import pandas as pd

def extract_machine_mapping(file_path, sheet_name):
    print(f"\n--- Extracting from {file_path} (Sheet: {sheet_name}) ---")
    try:
        # Read the file
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        # Find the header row (look for "EPR PARAMETER" or "No. Document" or "MC")
        header_idx = -1
        for i in range(min(15, len(df))):
            row_str = " ".join([str(x).upper() for x in df.iloc[i].values])
            if "EPR PARAMETER" in row_str or "MC" in row_str or "MACHINE" in row_str:
                header_idx = i
                break
                
        if header_idx == -1:
            print("Header not found.")
            return
            
        # Use the found row as header
        df.columns = [str(c).strip().upper() for c in df.iloc[header_idx]]
        df = df.iloc[header_idx+1:].reset_index(drop=True)
        
        # Find relevant columns
        doc_cols = [c for c in df.columns if "EPR PARAMETER" in c or "DOCUMENT" in c or "DOC" in c]
        mc_cols = [c for c in df.columns if c == "MC" or "MACHINE" in c]
        
        if not doc_cols or not mc_cols:
            print(f"Could not find required columns. Docs: {doc_cols}, MCs: {mc_cols}")
            print(f"Columns available: {list(df.columns)}")
            return
            
        doc_col = doc_cols[0]
        mc_col = mc_cols[0]
        print(f"Using Document Column: '{doc_col}', Machine Column: '{mc_col}'")
        
        # Print mappings
        count = 0
        for idx, row in df.iterrows():
            doc = str(row[doc_col]).strip()
            mc = str(row[mc_col]).strip()
            if doc and mc and doc != "nan" and mc != "nan" and doc != "-":
                print(f"  {doc:<35} -> {mc}")
                count += 1
        print(f"Total mappings found: {count}")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error: {e}")

# 1. Masterlist SPS Double Layer
extract_machine_mapping(r"d:\Project\extrude\folder excel\Masterlist SPS Double Layer.xlsx", "Parameter Setting")

# 2. Combined
extract_machine_mapping(r"d:\Project\extrude\folder excel\Masterlist_Digitalisasi_Combined_Fix.xlsx", "DL")
