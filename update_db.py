import pandas as pd
import json
import os

def update_db():
    try:
        df = pd.read_excel('Book1_tmp.xlsx')
        # Clean data: drop rows where Project is empty
        df = df.dropna(subset=['Project'])
        
        data = []
        for _, r in df.iterrows():
            # Determine status: default to IN PROGRESS unless mentioned otherwise
            # Check if there's a status column or if we should guess
            status = "IN PROGRESS"
            if 'Status' in df.columns and pd.notna(r['Status']):
                status = r['Status']
            
            data.append({
                "Customer": str(r['Project']).strip(),
                "Device": str(r['Solution Type']).strip() if pd.notna(r['Solution Type']) else "",
                "GW": str(r['GW']).strip() if pd.notna(r['GW']) else "",
                "Integrator": str(r['Integrator']).strip() if pd.notna(r['Integrator']) else "",
                "PM": str(r['Project Manager']).strip() if pd.notna(r['Project Manager']) else "Unassigned",
                "Version": str(r['Version']).strip() if pd.notna(r['Version']) else "",
                "Volume": str(r['Volume']).strip() if pd.notna(r['Volume']) else "",
                "Status": status
            })
            
        with open('integrations_db.json', 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        print(f"Successfully updated integrations_db.json with {len(data)} records.")
    except Exception as e:
        print(f"Error during update: {e}")

if __name__ == "__main__":
    update_db()
