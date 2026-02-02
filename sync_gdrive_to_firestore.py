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
    
    # 1. Load existing categories from Firestore to avoid duplicates
    existing_cats = {}
    try:
        docs = db.collection('guides_categories').stream()
        for doc in docs:
            d = doc.to_dict()
            existing_cats[d['name']] = d['id']
    except Exception as e:
        print(f"Error loading existing categories: {e}")

    root_files = list_files_in_folder(drive, GDRIVE_FOLDER_ID)
    print(f"Found {len(root_files)} main items in Root.")

    for root_f in root_files:
        name = root_f['name']
        fid = root_f['id']
        mime = root_f['mimeType']
        
        print(f"\n[GDrive Item] {name} ({fid})")
        
        if mime == 'application/vnd.google-apps.folder':
            # This is a Category folder
            cat_id = existing_cats.get(name)
            if not cat_id:
                # Create new category
                cat_id = str(uuid.uuid4())
                print(f"   Creating NEW Category: {name} ({cat_id})")
                db.collection('guides_categories').document(cat_id).set({
                    "id": cat_id,
                    "name": name,
                    "emoji": "📂",
                    "type": "kb",
                    "guides": [],
                    "subCategories": []
                })
            
            sub_files = list_files_in_folder(drive, fid)
            print(f"   Syncing {len(sub_files)} files into category '{name}'")
            
            for f in sub_files:
                if f['mimeType'] == 'application/vnd.google-apps.folder':
                    # Treat subfolders as Sub-Categories
                    sub_sub_files = list_files_in_folder(drive, f['id'])
                    print(f"      Sub-Category: {f['name']} ({len(sub_sub_files)} items)")
                    for ssf in sub_sub_files:
                        save_to_firestore(db, ssf, cat_id, sub_category=f['name'])
                else:
                    save_to_firestore(db, f, cat_id)
        else:
            # Root file (not in folder) - Put in a "General" category
            gen_cat_name = "כללי"
            cat_id = existing_cats.get(gen_cat_name)
            if not cat_id:
                cat_id = "general"
                db.collection('guides_categories').document(cat_id).set({
                    "id": cat_id, "name": gen_cat_name, "emoji": "📁", "type": "kb", "guides": [], "subCategories": []
                })
            save_to_firestore(db, root_f, cat_id)

    print("\nSync Complete!")

def save_to_firestore(db, file_data, cat_id, sub_category=None):
    if file_data['mimeType'] == 'application/vnd.google-apps.folder': return

    name = file_data['name']
    # Use direct download link or webViewLink
    url = file_data.get('webViewLink')
    # If we want direct download for images/PDFs:
    # url = f"https://drive.google.com/uc?export=download&id={file_data['id']}"
    
    # We add the guide to the 'guides' collection, linked by Category ID
    gid = str(uuid.uuid5(uuid.NAMESPACE_DNS, name + (sub_category or "") + cat_id))
    
    print(f"      Saving: {name}")
    doc_ref = db.collection('guides').document(gid)
    doc_ref.set({
        "id": gid,
        "title": name,
        "content": f"קובץ מגוגל דרייב: {name}\n\nקישור לצפייה: {url}",
        "Category": cat_id,
        "SubCategory": sub_category,
        "url": url,
        "gdrive_id": file_data['id'],
        "type": "file",
        "last_sync": firestore.SERVER_TIMESTAMP
    }, merge=True)

if __name__ == "__main__":
    sync()
