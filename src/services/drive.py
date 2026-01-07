import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from src.utils.config import settings
from src.utils.logger import logger
from typing import Optional, List, Dict

class DriveService:
    SCOPES = ['https://www.googleapis.com/auth/drive']

    def __init__(self):
        self.service = None
        self._connect()

    def _connect(self):
        """Connect to Google Drive API."""
        try:
            creds = Credentials.from_service_account_file(
                settings.google_credentials_file, scopes=self.SCOPES
            )
            self.service = build('drive', 'v3', credentials=creds)
            logger.info("connected_to_google_drive")
        except Exception as e:
            logger.error("google_drive_connection_failed", error=str(e))
            raise

    def create_folder(self, name: str, parent_id: Optional[str] = None) -> str:
        """Create a folder and return its ID."""
        try:
            file_metadata = {
                'name': name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            if parent_id:
                file_metadata['parents'] = [parent_id]

            file = self.service.files().create(body=file_metadata, fields='id').execute()
            logger.info("folder_created", name=name, id=file.get('id'))
            return file.get('id')
        except Exception as e:
            logger.error("create_folder_failed", name=name, error=str(e))
            raise

    def copy_file(self, file_id: str, dest_folder_id: str, new_name: Optional[str] = None) -> str:
        """Copy a file to a destination folder."""
        try:
            # Get original name if not provided
            if not new_name:
                orig = self.service.files().get(fileId=file_id, fields='name').execute()
                new_name = orig.get('name')

            file_metadata = {
                'name': new_name,
                'parents': [dest_folder_id]
            }

            file = self.service.files().copy(
                fileId=file_id, 
                body=file_metadata, 
                fields='id'
            ).execute()
            
            logger.info("file_copied", original_id=file_id, new_id=file.get('id'))
            return file.get('id')
        except Exception as e:
            logger.error("copy_file_failed", file_id=file_id, error=str(e))
            raise

    def get_file_id_from_url(self, url: str) -> Optional[str]:
        """Extract file ID from a Google Drive URL."""
        if not url:
            return None
            
        import re
        # Match /d/ID or id=ID patterns
        match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
        if match:
            return match.group(1)
        
        match = re.search(r'id=([a-zA-Z0-9_-]+)', url)
        if match:
            return match.group(1)
            
        return None

# Singleton
_drive_service: Optional[DriveService] = None

def get_drive_service() -> DriveService:
    global _drive_service
    if _drive_service is None:
        _drive_service = DriveService()
    return _drive_service
