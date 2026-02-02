import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import os
import json

# --- CONFIG ---
EXCEL_FILE = '_לקוחות ורטיקלים - אחריות.xlsx'
FIREBASE_CREDS_FILE = 'tier-2-vico-firebase-adminsdk.json'

def init_firestore():
    if not firebase_admin._apps:
        cred = credentials.Certificate(FIREBASE_CREDS_FILE)
        firebase_admin.initialize_app(cred)
    return firestore.client()

def sync_warranty():
    print("Reading Excel...")
    df = pd.read_excel(EXCEL_FILE)
    
    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]
    
    db = init_firestore()
    doc_ref = db.collection('data').document('integrations')
    doc = doc_ref.get()
    
    if not doc.exists:
        print("Integrations document not found in Firestore. Reading local fallback...")
        p = os.path.join(os.path.dirname(__file__), 'integrations_db.json')
        if os.path.exists(p):
            with open(p, 'r', encoding='utf-8') as f:
                integrations = json.load(f)
        else:
            print("Local integrations_db.json also missing!")
            integrations = []
    else:
        integrations = doc.to_dict().get('list', [])
        
    print(f"Loaded {len(integrations)} integrations from base.")
    
    updated_count = 0
    for idx, row in df.iterrows():
        customer_name = str(row.get('חברה', '')).strip()
        if not customer_name: continue
        
        # Look for match in integrations
        match = None
        for i in integrations:
            if i['Customer'].lower() == customer_name.lower():
                match = i
                break
        
        # If no match, we might need a new integration entry or skip
        if not match:
            # For this task, let's create a new one if it's a major customer
            match = {
                "id": str(idx + 1000),
                "Customer": customer_name,
                "ProjectManager": "Unknown",
                "Status": "Active"
            }
            integrations.append(match)
            print(f" + Added new customer: {customer_name}")

        # Update warranty fields
        match['WarrantyStatus'] = str(row.get('אחריות נכון ל 2026', 'n/a')).strip()
        match['WarrantyDuration'] = str(row.get('משך האחריות בזמן עסקה', '')).strip()
        match['ServiceResponse'] = str(row.get('מענה שירות לקוחות', '')).strip()
        match['WarrantyCoverage'] = str(row.get('מה כוללת האחריות', '')).strip()
        match['EquipmentRepair'] = str(row.get('איך מגיע ציוד לתיקון ושליחת טכנאי', '')).strip()
        match['SLA'] = str(row.get('SLA', '')).strip()
        
        updated_count += 1

    # Save back to Firestore
    doc_ref.set({"list": integrations}, merge=True)
    
    # Save back to local JSON for fallback consistency
    p = os.path.join(os.path.dirname(__file__), 'integrations_db.json')
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(integrations, f, indent=4, ensure_ascii=False)
        
    print(f"Successfully updated {updated_count} customers with warranty data (Firestore & local JSON).")

if __name__ == "__main__":
    sync_warranty()
