import win32com.client
import os
import pandas as pd
import tempfile
import re

def clean_phone(p):
    s = str(p).strip()
    if s.endswith('.0'): s = s[:-2]
    return re.sub(r'\D', '', s)

def scan_folder_recursive(folder, subjects_to_find, found_files, depth=0):
    if depth > 3: return # Limit recursion
    
    try:
        # Check items in current folder
        # Optimization: Restrict to last 30 days or just scan last 50 items
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        
        count = 0
        for item in items:
            count += 1
            if count > 50: break # Only check recent emails
            
            try:
                subj = item.Subject
                for s in subjects_to_find:
                    if s.lower() in subj.lower():
                        print(f"FAILED TO FIND in Inbox, but FOUND in '{folder.Name}': '{subj}'")
                        if item.Attachments.Count > 0:
                            for att in item.Attachments:
                                fname = att.FileName.lower()
                                if fname.endswith(('.csv', '.xls', '.xlsx')):
                                    temp_dir = tempfile.gettempdir()
                                    fpath = os.path.join(temp_dir, f"Debug_{att.FileName}")
                                    try:
                                        att.SaveAsFile(fpath)
                                        found_files.append(fpath)
                                        print(f"  Downloaded: {fpath}")
                                    except: pass
            except: pass
        
        # Recurse
        for sub in folder.Folders:
            scan_folder_recursive(sub, subjects_to_find, found_files, depth+1)
            
    except Exception as e:
        # print(f"Error scanning {folder.Name}: {e}")
        pass

def scan_outlook_for_subjects():
    print("Connecting to Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    
    subjects_to_find = [
        "Log_VICO_Monthly",
        "Scheduled Report: Survey Result_new_MONTHLY"
    ]
    
    print(f"Scanning recursively for subjects: {subjects_to_find}")
    found_files = []
    
    scan_folder_recursive(inbox, subjects_to_find, found_files)
    
    return found_files

def analyze_file(fpath):
    print(f"\nAnalyzing: {os.path.basename(fpath)}")
    try:
        if fpath.endswith('.csv'):
            try: df = pd.read_csv(fpath)
            except: df = pd.read_csv(fpath, encoding='cp1255')
        else:
            df = pd.read_excel(fpath)
            
        print(f"  Columns: {list(df.columns)}")
        
        dnis_col = next((c for c in df.columns if 'DNIS' in c.upper() or 'DIALED TO' in c.upper()), None)
        campaign_col = next((c for c in df.columns if 'CAMPAIGN' in c.upper()), None)
        
        if dnis_col:
            print(f"  Found DNIS Column: {dnis_col}")
            df['CleanDNIS'] = df[dnis_col].apply(clean_phone)
            counts = df['CleanDNIS'].value_counts().head(20)
            print("  Top 20 DNIS/Phone numbers found:")
            print(counts)
            
        if campaign_col:
            print(f"  Found Campaign Column: {campaign_col}")
            counts = df[campaign_col].value_counts().head(20)
            print("  Top 20 Campaigns found:")
            print(counts)
            
    except Exception as e:
        print(f"  Error analyzing file: {e}")

if __name__ == "__main__":
    files = scan_outlook_for_subjects()
    if not files:
        print("\nNo matching emails found recursively.")
    else:
        for f in files:
            analyze_file(f)
