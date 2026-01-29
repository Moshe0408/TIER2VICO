"""
Firebase Data Migration Script - Firestore Only
Migrates guides_db.json to Firestore
Files remain in local uploads/ directory
"""

import firebase_admin
from firebase_admin import credentials, firestore
import json
import os
from pathlib import Path

# Initialize Firebase Admin SDK
cred = credentials.Certificate("tier-2-vico-firebase-adminsdk.json")
# Try to auto-discover or use common patterns
firebase_admin.initialize_app(cred)

db = firestore.client()
BASE_DIR = Path(__file__).parent

from firebase_admin import storage

def upload_to_firebase(file_path):
    # Potential bucket names to try
    buckets_to_try = [
        "tier-2-vico.appspot.com",
        "tier-2-vico.firebasestorage.app"
    ]
    
    for bname in buckets_to_try:
        try:
            bucket = storage.bucket(bname)
            blob = bucket.blob(f"uploads/{file_path.name}")
            blob.upload_from_filename(str(file_path))
            blob.make_public()
            return blob.public_url
        except Exception:
            continue
            
    # Fallback: Return local path if cloud upload fails
    # This allows indexing the metadata even if files are local
    print(f"Cloud upload failed for {file_path.name}. Using local path fallback.")
    return f"/לקוחות/{file_path.parent.name}/{file_path.name}"

def migrate():
    """Migrate guides and upload their files to Drive"""
    guides_file = BASE_DIR / "guides_db.json"
    if not guides_file.exists(): return
    
    with open(guides_file, 'r', encoding='utf-8-sig') as f:
        guides_data = json.load(f)
    
    print(f"Processing {len(guides_data)} guides...")
    uploads_map = {} # local_path -> drive_url
    
    for guide in guides_data:
        # Check for local upload paths in guide and sub-categories
        def process_item(item):
            # Check for strings like "/uploads/..."
            item_str = json.dumps(item)
            local_paths = re.findall(r'/uploads/[a-zA-Z0-9\-\.]+', item_str)
            for lp in local_paths:
                if lp not in uploads_map:
                    local_file = BASE_DIR / lp.lstrip('/')
                    if local_file.exists():
                        print(f"Uploading {lp} to Drive...")
                        drive_url = upload_to_firebase(local_file)
                        if drive_url:
                            uploads_map[lp] = drive_url
                            print(f"  ✓ {lp} -> {drive_url}")
            
            # Replace paths in the item
            new_item_str = item_str
            for lp, du in uploads_map.items():
                new_item_str = new_item_str.replace(lp, du)
            return json.loads(new_item_str)

        migrated_guide = process_item(guide)
        doc_id = migrated_guide.get('id', str(uuid.uuid4()))
        db.collection('guides').document(doc_id).set(migrated_guide)
        print(f"✓ Migrated Guide: {migrated_guide.get('name')}")

def migrate_release_docs():
    """Scan 'לקוחות' for release documents and upload to Drive"""
    cust_dir = BASE_DIR / "לקוחות"
    if not cust_dir.exists():
        print("לקוחות directory not found.")
        return

    print(f"Scanning {cust_dir} for release documents...")
    
    # We look for "טופס סיום פיילוט" or similar PROD forms
    found_files = []
    for root, dirs, files in os.walk(cust_dir):
        for file in files:
            if "טופס סיום" in file or "PROD" in file or "יציאה לייצור" in file:
                found_files.append(Path(root) / file)
    
    if not found_files:
        print("No release documents found.")
        return

    print(f"Found {len(found_files)} release documents. Processing...")
    
    # Check if 'מרכז ידע' category exists or create it
    # For now, let's put them in a dedicated 'Release Documents' category in Firestore
    cat_id = "release-docs-cat"
    category = {
        "id": cat_id,
        "name": "ישיבות שחרור וטפסים",
        "emoji": "📄",
        "type": "kb",
        "subCategories": []
    }
    
    guides = []
    for f in found_files:
        print(f"Uploading {f.name}...")
        url = upload_to_firebase(f)
        if url:
            guide = {
                "id": str(uuid.uuid4()),
                "name": f.name,
                "content": f"טופס שחרור עבור {f.parent.name}",
                "attachments": [url],
                "Category": cat_id,
                "date": datetime.now().strftime("%Y-%m-%d")
            }
            guides.append(guide)
            
            # Create a separate sub-category per customer if needed
            subcat_name = f.parent.name
            if not any(s['name'] == subcat_name for s in category['subCategories']):
                category['subCategories'].append({
                    "id": str(uuid.uuid4()),
                    "name": subcat_name
                })
    
    # Save the category
    db.collection('guides_categories').document(cat_id).set(category)
    
    # Save the guides
    for g in guides:
        db.collection('guides').document(g['id']).set(g)
        print(f"  ✓ {g['name']} -> Firestore")

if __name__ == "__main__":
    import re, uuid
    from datetime import datetime
    migrate()
    migrate_release_docs()

