import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build

GDRIVE_CREDS_FILE = 'google-drive-credentials.json'
# ID of מדריכים from previous run
FOLDER_ID = "1Oy7zU9AW5arwvECskavtYcD2Uq4_YHBI"

def init_gdrive():
    creds = service_account.Credentials.from_service_account_file(
        GDRIVE_CREDS_FILE, 
        scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=creds)

def list_all(service, folder_id):
    print(f"Listing everything in folder: {folder_id}")
    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(
        q=query, 
        fields="files(id, name, mimeType)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = results.get('files', [])
    if not files:
        print("Folder is EMPTY.")
    for f in files:
        print(f" - {f['name']} ({f['id']}) [{f['mimeType']}]")

if __name__ == "__main__":
    service = init_gdrive()
    list_all(service, FOLDER_ID)
