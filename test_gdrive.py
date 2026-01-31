import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload
from google.oauth2 import service_account

try:
    creds = service_account.Credentials.from_service_account_file(
        'google-drive-credentials.json',
        scopes=['https://www.googleapis.com/auth/drive']
    )
    service = build('drive', 'v3', credentials=creds)
    file_metadata = {'name': 'test.txt'}
    media = MediaInMemoryUpload(b'test', mimetype='text/plain')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f'Success: {file.get("id")}')
except Exception as e:
    print(f'Error: {e}')
