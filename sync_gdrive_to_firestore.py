import os
import json
import uuid
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
import firebase_admin
from firebase_admin import credentials, firestore

# --- CONFIG ---
GDRIVE_CREDS_FILE = 'google-drive-credentials.json'
GDRIVE_FOLDER_ID = "13jBR4dJOhojtf63_mGYLeoAqiqP7KjJs"
FIREBASE_CREDS_FILE = 'tier-2-vico-firebase-adminsdk.json'

# --- INIT ---
def init_gdrive():
    creds = service_account.Credentials.from_service_account_file(
        GDRIVE_CREDS_FILE, 
        scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=creds)

def init_firestore():
    if not firebase_admin._apps:
        cred = credentials.Certificate(FIREBASE_CREDS_FILE)
        firebase_admin.initialize_app(cred)
    return firestore.client()

def list_files_in_folder(service, folder_id):
    files = []
    page_token = None
    while True:
        try:
            query = f"'{folder_id}' in parents and trashed = false"
            results = service.files().list(
                q=query, 
                fields="nextPageToken, files(id, name, mimeType, webViewLink, webContentLink)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                pageSize=1000,
                pageToken=page_token
            ).execute()
            files.extend(results.get('files', []))
            page_token = results.get('nextPageToken')
            if not page_token:
                break
        except Exception as e:
            print(f"Error listing folder {folder_id}: {e}")
            break
    return files

def sync():
    print("Connecting to APIs...")
    drive = init_gdrive()
    db = init_firestore()
    
    root_files = list_files_in_folder(drive, GDRIVE_FOLDER_ID)
    print(f"Found {len(root_files)} main items in Root.")

    for root_f in root_files:
        name = root_f['name']
        fid = root_f['id']
        mime = root_f['mimeType']
        
        print(f"\n[Root Item] {name} ({fid})")
        
        if mime == 'application/vnd.google-apps.folder':
            # Map top-level Hebrew folders to App Categories
            app_cat = "kb-guides"
            if "לקוחות" in name: app_cat = "integrations"
            elif "דרייבר" in name or "דרייבירם" in name: app_cat = "kb-drivers"
            
            print(f"   Category Mapping: {app_cat}")
            
            sub_files = list_files_in_folder(drive, fid)
            print(f"   Found {len(sub_files)} items in {name}")
            
            for f in sub_files:
                if f['mimeType'] == 'application/vnd.google-apps.folder':
                    print(f"      Scanning Subfolder: {f['name']}")
                    sub_sub_files = list_files_in_folder(drive, f['id'])
                    for ssf in sub_sub_files:
                        save_to_firestore(db, ssf, app_cat, sub_category=f['name'])
                else:
                    save_to_firestore(db, f, app_cat)
        else:
            save_to_firestore(db, root_f, "others")

    print("\nSync Complete!")

def save_to_firestore(db, file_data, category, sub_category=None):
    if file_data['mimeType'] == 'application/vnd.google-apps.folder': return

    name = file_data['name']
    url = f"https://drive.google.com/uc?export=download&id={file_data['id']}"
    
    # Clean category/sub_category for Firestore IDs if needed
    
    # For Guides (kb-guides, kb-drivers)
    if category.startswith("kb"):
        # Use filename as ID to prevent duplicates on re-run
        gid = str(uuid.uuid5(uuid.NAMESPACE_DNS, name + (sub_category or "")))
        doc_ref = db.collection('guides').document(gid)
        doc_ref.set({
            "id": gid,
            "title": name,
            "content": f"Download Link: {url}",
            "Category": sub_category or category,
            "url": url,
            "gdrive_id": file_data['id'],
            "type": "file",
            "last_sync": firestore.SERVER_TIMESTAMP
        }, merge=True)
    
    # For Integrations (Customer documents)
    elif category == "integrations" and sub_category:
        docs = db.collection('data').document('integrations').get()
        if docs.exists:
            ints = docs.to_dict().get('list', [])
            updated = False
            for i in ints:
                # Fuzzy match customer name
                if i['Customer'].lower() in sub_category.lower() or sub_category.lower() in i['Customer'].lower():
                    if "Sheet" in name: i['SheetURL'] = url
                    elif "Note" in name: i['NoteURL'] = url
                    else: i['ManualURL'] = url
                    updated = True
                    print(f"         Matched Project: {i['Customer']}")
            if updated:
                db.collection('data').document('integrations').set({"list": ints}, merge=True)

if __name__ == "__main__":
    sync()
