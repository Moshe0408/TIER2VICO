import pandas as pd
import json
import os

def sync_data():
    try:
        # Read the file
        df = pd.read_excel('Book1_sync.xlsx')
        
        # Drop rows where Project is empty
        df = df.dropna(subset=['Project'])
        
        # Remove duplicates based on Project name
        df = df.drop_duplicates(subset=['Project'], keep='first')
        
        data = []
        for _, r in df.iterrows():
            # Clean version string: remove "Terminal:" if present
            version = str(r['Version']).replace('Terminal:', '').strip() if pd.notna(r['Version']) else ""
            
            data.append({
                "Customer": str(r['Project']).strip(),
                "Device": str(r['Solution Type']).strip() if pd.notna(r['Solution Type']) else "",
                "GW": str(r['GW']).strip() if pd.notna(r['GW']) else "",
                "Integrator": str(r['Integrator']).strip() if pd.notna(r['Integrator']) else "",
                "PM": str(r['Project Manager']).strip() if pd.notna(r['Project Manager']) else "Unassigned",
                "Version": version,
                "Status": "IN PROGRESS" # Default status
            })
            
        with open('integrations_db.json', 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
            
        print(f"Successfully synced {len(data)} unique projects to integrations_db.json")
    except Exception as e:
        print(f"Sync error: {e}")

if __name__ == "__main__":
    sync_data()
