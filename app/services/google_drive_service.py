import os
import json
import logging
from io import BytesIO
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Scopes required for Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive.file']

logger = logging.getLogger(__name__)

class GoogleDriveService:
    def __init__(self):
        self.folder_id = os.getenv("GOOGLE_DRIVE_FOLDER_ID")
        self.credentials_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
        self.service = self._authenticate()

    def _authenticate(self):
        """Authenticates using service account credentials from env or file."""
        try:
            if self.credentials_json:
                # If credentials_json starts with '{', it's likely the raw JSON string
                if self.credentials_json.strip().startswith('{'):
                    info = json.loads(self.credentials_json)
                    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
                else:
                    # Otherwise, treat it as a file path
                    creds = service_account.Credentials.from_service_account_file(self.credentials_json, scopes=SCOPES)
            else:
                # Fallback to default file name if not in env
                creds_path = "service_account.json"
                if os.path.exists(creds_path):
                    creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
                else:
                    logger.warning("No Google Drive credentials found. Upload will fail.")
                    return None
            
            return build('drive', 'v3', credentials=creds)
        except Exception as e:
            logger.error(f"Failed to authenticate with Google Drive: {e}")
            return None

    def upload_file(self, file_stream: BytesIO, filename: str) -> dict:
        """Uploads a file stream to the specified Google Drive folder."""
        if not self.service:
            return {"success": False, "error": "Google Drive service not authenticated"}

        try:
            file_metadata = {
                'name': filename,
                'parents': [self.folder_id] if self.folder_id else []
            }
            
            # Reset stream position to beginning
            file_stream.seek(0)
            
            media = MediaIoBaseUpload(
                file_stream, 
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                resumable=True
            )
            
            file = self.service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, webViewLink'
            ).execute()
            
            return {
                "success": True,
                "file_id": file.get('id'),
                "view_link": file.get('webViewLink')
            }
            
        except Exception as e:
            logger.error(f"Error uploading file to Google Drive: {e}")
            return {"success": False, "error": str(e)}

# Singleton instance
drive_service = GoogleDriveService()
