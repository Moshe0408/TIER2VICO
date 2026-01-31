import firebase_admin
from firebase_admin import credentials, storage
import json
from google.cloud import storage as gcs

cred_file = "tier-2-vico-firebase-adminsdk.json"
with open(cred_file) as f:
    cred_data = json.load(f)

project_id = cred_data.get('project_id')
print(f"Project ID: {project_id}")

try:
    client = gcs.Client.from_service_account_json(cred_file)
    buckets = list(client.list_buckets())
    print("Available Buckets:")
    for b in buckets:
        print(f" - {b.name}")
except Exception as e:
    print(f"Error listing buckets: {e}")
