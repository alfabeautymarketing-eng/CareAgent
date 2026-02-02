from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from src.utils.config import settings
from src.utils.logger import logger
from typing import Optional, List, Dict, Any

class DocsService:
    SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']

    def __init__(self):
        self.service = None
        self._connect()

    def _connect(self):
        """Connect to Google Docs API."""
        try:
            creds = Credentials.from_service_account_file(
                settings.google_credentials_file, scopes=self.SCOPES
            )
            self.service = build('docs', 'v1', credentials=creds)
            logger.info("connected_to_google_docs_api")
        except Exception as e:
            logger.error("google_docs_connection_failed", error=str(e))
            raise

    def create_document(self, title: str) -> str:
        """Create a new Google Doc and return its ID."""
        try:
            body = {'title': title}
            doc = self.service.documents().create(body=body).execute()
            logger.info("document_created", title=title, id=doc.get('documentId'))
            return doc.get('documentId')
        except Exception as e:
            logger.error("create_document_failed", title=title, error=str(e))
            raise

    def batch_update(self, document_id: str, requests: List[Dict[str, Any]]):
        """Execute a batch update on a document."""
        try:
            self.service.documents().batchUpdate(
                documentId=document_id,
                body={'requests': requests}
            ).execute()
        except Exception as e:
            logger.error("docs_batch_update_failed", id=document_id, error=str(e))
            raise

    def clear_document(self, document_id: str):
        """Clear document content (effectively)."""
        # Google Docs API doesn't have a single 'clear' command.
        # We get the end index and delete everything from 1 to end-1.
        try:
            doc = self.service.documents().get(documentId=document_id).execute()
            content = doc.get('body').get('content')
            last_element = content[-1]
            end_index = last_element.get('endIndex')
            
            if end_index > 2:
                self.batch_update(document_id, [
                    {
                        'deleteContentRange': {
                            'range': {
                                'startIndex': 1,
                                'endIndex': end_index - 1
                            }
                        }
                    }
                ])
        except Exception as e:
            logger.error("clear_document_failed", id=document_id, error=str(e))
            raise

    def append_text(self, document_id: str, text: str, is_heading: bool = False):
        """Append text to document."""
        requests = [
            {
                'insertText': {
                    'location': {'index': self._get_end_index(document_id)},
                    'text': text + "\n"
                }
            }
        ]
        
        # If it was a heading, we'd need to apply paragraph style to the newly inserted range.
        # But indices change, so it's easier to do it in one batch.
        self.batch_update(document_id, requests)

    def _get_end_index(self, document_id: str) -> int:
        doc = self.service.documents().get(documentId=document_id).execute()
        return doc.get('body').get('content')[-1].get('endIndex') - 1

# Singleton
_docs_service: Optional[DocsService] = None

def get_docs_service() -> DocsService:
    global _docs_service
    if _docs_service is None:
        _docs_service = DocsService()
    return _docs_service
