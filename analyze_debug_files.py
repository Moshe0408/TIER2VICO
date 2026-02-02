import pandas as pd
import os
import re

def clean_phone(p):
    s = str(p).strip()
    if s.endswith('.0'): s = s[:-2]
    return re.sub(r'\D', '', s)

def analyze_smart(fpath):
    print(f"\nAnalyzing Smart: {os.path.basename(fpath)}")
    try:
        raw_df = pd.read_excel(fpath, header=None)
        header_row = 0
        found = False
        for i, row in raw_df.iterrows():
            row_str = " ".join([str(val).upper() for val in row if pd.notna(val)])
            if 'AGENT NAME' in row_str or 'TALK TIME' in row_str or 'HANDLE TIME' in row_str or 'DNIS' in row_str or 'CAMPAIGN' in row_str:
                header_row = i
                found = True
                break
        
        if found:
            print(f"  Header found at row {header_row}")
            df = pd.read_excel(fpath, skiprows=header_row)
        else:
            print("  Header NOT found, using default.")
            df = pd.read_excel(fpath)

        df.columns = [str(c).strip().upper() for c in df.columns]
        print(f"  Columns: {list(df.columns)}")

        dnis_col = next((c for c in df.columns if 'DNIS' in c or 'DIALED TO' in c), None)
        campaign_col = next((c for c in df.columns if 'CAMPAIGN' in c), None)
        
        if dnis_col:
            print(f"  Found DNIS Column: {dnis_col}")
            df['CleanDNIS'] = df[dnis_col].apply(clean_phone)
            counts = df['CleanDNIS'].value_counts()
            print("  DNIS counts:")
            print(counts)
            
        if campaign_col:
            print(f"  Found Campaign Column: {campaign_col}")
            counts = df[campaign_col].value_counts()
            print("  Campaign counts:")
            print(counts)

    except Exception as e:
        print(f"Error: {e}")

f1 = r"C:\Users\moshei1\AppData\Local\Temp\Debug_Call Log_VICO_Monthly_Moshe.xlsx"
if os.path.exists(f1): analyze_smart(f1)

f2 = r"C:\Users\moshei1\AppData\Local\Temp\Debug_Survey Result_new_MONTHLY.xlsx"
if os.path.exists(f2): analyze_smart(f2)
