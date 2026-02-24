import os
import json
import logging
from io import BytesIO
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Scopes required for Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']

logger = logging.getLogger(__name__)

class GoogleDriveService:
    def __init__(self):
        self.folder_id = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "").strip()
        self.client_id = os.getenv("GOOGLE_CLIENT_ID", "").strip()
        self.client_secret = os.getenv("GOOGLE_CLIENT_SECRET", "").strip()
        self.refresh_token = os.getenv("GOOGLE_REFRESH_TOKEN", "").strip()
        
        # Log limited info for debugging in Render
        if self.folder_id:
            logger.info(f"GoogleDriveService initialized with folder_id: {self.folder_id[:5]}...{self.folder_id[-5:]}")
        else:
            logger.warning("GoogleDriveService initialized WITHOUT GOOGLE_DRIVE_FOLDER_ID")
            
        self.service = self._authenticate()

    def _authenticate(self):
        """Authenticates using OAuth2 Client ID, Secret, and Refresh Token."""
        try:
            if not (self.client_id and self.client_secret and self.refresh_token):
                logger.warning("Missing OAuth2 credentials (ID, Secret, or Token). Upload will fail.")
                return None
            
            creds = Credentials(
                token=None,  # Will be refreshed
                refresh_token=self.refresh_token,
                client_id=self.client_id,
                client_secret=self.client_secret,
                token_uri="https://oauth2.googleapis.com/token",
                scopes=SCOPES
            )
            
            # Force refresh to ensure token is valid
            creds.refresh(Request())
            
            return build('drive', 'v3', credentials=creds)
        except Exception as e:
            logger.error(f"Failed to authenticate with Google Drive OAuth2: {e}")
            return None

    def upload_file(self, file_stream: BytesIO, filename: str) -> dict:
        """Uploads a file stream to the specified Google Drive folder."""
        if not self.service:
            return {"success": False, "error": "Google Drive service not authenticated (check OAuth2 credentials)"}

        try:
            logger.info(f"Uploading file to Drive: {filename}")
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
                fields='id, webViewLink, webContentLink'
            ).execute()
            
            view_link = file.get('webViewLink')
            download_link = file.get('webContentLink')
            file_id = file.get('id')
            logger.info(f"Successfully uploaded to Drive. ID: {file_id}")
            logger.info(f"View Link: {view_link}")
            logger.info(f"Download Link: {download_link}")
            
            return {
                "success": True,
                "file_id": file_id,
                "view_link": view_link,
                "download_link": download_link
            }
            
        except Exception as e:
            logger.error(f"Error uploading file to Google Drive: {e}")
            return {"success": False, "error": str(e)}

# Singleton instance
drive_service = GoogleDriveService()
