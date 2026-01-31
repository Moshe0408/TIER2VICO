"""
Final Migration Script - Google Drive Storage + Firestore
Migrates guides and release documents from local 'לקוחות' folder to Google Drive.
Requires shared folder access.
"""

import os
import json
import re
import uuid
from pathlib import Path
from datetime import datetime
import firebase_admin
from firebase_admin import credentials, firestore
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Configuration
DRIVE_CREDENTIALS_FILE = "google-drive-credentials.json"
FIREBASE_CREDENTIALS_FILE = "tier-2-vico-firebase-adminsdk.json"
SHARED_FOLDER_ID = "13jBR4dJOhojtf63_mGYLeoAqiqP7KjJs"

BASE_DIR = Path(__file__).parent

def get_drive_service():
    creds = service_account.Credentials.from_service_account_file(
        DRIVE_CREDENTIALS_FILE, 
        scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

def upload_to_drive(file_path):
    """Uploads a file to the shared Google Drive folder and returns the view URL"""
    service = get_drive_service()
    
    file_metadata = {
        'name': file_path.name,
        'parents': [SHARED_FOLDER_ID]
    }
    media = MediaFileUpload(str(file_path), resumable=True)
    
    try:
        # Note: Even if uploading to a shared folder, if the সার্ভিস অ্যাকাউন্ট isn't part of a Google Workspace/Shared Drive
        # it might still use its own quota. 
        # However, the user provided a personal folder link which should work if shared correctly.
        file = service.files().create(
            body=file_metadata, 
            media_body=media, 
            fields='id, webViewLink',
            supportsAllDrives=True # Important for shared folders/drives
        ).execute()
        return file.get('webViewLink')
    except Exception as e:
        print(f"Error uploading {file_path.name}: {e}")
        return None

# Initialize Firebase
if not firebase_admin._apps:
    cred = credentials.Certificate(FIREBASE_CREDENTIALS_FILE)
    firebase_admin.initialize_app(cred)

db = firestore.client()

def migrate_guides():
    """Migrate guides_db.json to Firestore and move local uploads to Drive"""
    guides_file = BASE_DIR / "guides_db.json"
    if not guides_file.exists(): 
        print("guides_db.json not found.")
        return
    
    with open(guides_file, 'r', encoding='utf-8-sig') as f:
        guides_data = json.load(f)
    
    print(f"Processing {len(guides_data)} guides...")
    uploads_map = {}
    
    for guide in guides_data:
        # Scan for local /uploads/ paths
        def process_item(item):
            item_str = json.dumps(item)
            local_paths = re.findall(r'/uploads/[a-zA-Z0-9\-\.]+', item_str)
            for lp in local_paths:
                if lp not in uploads_map:
                    local_file = BASE_DIR / lp.lstrip('/')
                    if local_file.exists():
                        print(f"Uploading {lp} to Drive...")
                        drive_url = upload_to_drive(local_file)
                        if drive_url:
                            uploads_map[lp] = drive_url
                            print(f"  ✓ {lp} -> {drive_url}")
            
            new_item_str = item_str
            for lp, du in uploads_map.items():
                new_item_str = new_item_str.replace(lp, du)
            return json.loads(new_item_str)

        migrated_guide = process_item(guide)
        doc_id = migrated_guide.get('id', str(uuid.uuid4()))
        db.collection('guides').document(doc_id).set(migrated_guide)
        print(f"✓ Guide '{migrated_guide.get('name')}' migrated.")

def migrate_release_docs():
    """Scan 'לקוחות' for release documents and upload to Drive"""
    cust_dir = BASE_DIR / "לקוחות"
    if not cust_dir.exists():
        print("לקוחות directory not found.")
        return

    print("Scanning for release documents (טופס סיום / PROD)...")
    found_files = []
    for root, dirs, files in os.walk(cust_dir):
        for file in files:
            if any(kw in file for kw in ["טופס סיום", "PROD", "יציאה לייצור"]):
                found_files.append(Path(root) / file)
    
    if not found_files:
        print("No release documents found.")
        return

    cat_id = "release-docs-cat"
    category = {
        "id": cat_id,
        "name": "ישיבות שחרור וטפסים",
        "emoji": "📄",
        "type": "kb",
        "subCategories": []
    }
    
    for f in found_files:
        print(f"Uploading {f.name}...")
        url = upload_to_drive(f)
        if url:
            guide_id = str(uuid.uuid4())
            subcat_name = f.parent.name
            
            # Find or create subcategory
            subcat = next((s for s in category['subCategories'] if s['name'] == subcat_name), None)
            if not subcat:
                subcat = {"id": str(uuid.uuid4()), "name": subcat_name}
                category['subCategories'].append(subcat)

            guide = {
                "id": guide_id,
                "name": f.name,
                "content": f"טופס שחרור עבור {subcat_name}",
                "attachments": [url],
                "Category": cat_id,
                "subCategory": subcat['id'],
                "date": datetime.now().strftime("%Y-%m-%d")
            }
            db.collection('guides').document(guide_id).set(guide)
            print(f"  ✓ {f.name} -> Drive & Firestore")
    
    db.collection('guides_categories').document(cat_id).set(category)

if __name__ == "__main__":
    print("Starting Mega Migration...")
    migrate_guides()
    migrate_release_docs()
    print("Migration Complete.")
