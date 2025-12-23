"""
PDF Processor for AgentCare.
Extracts text from PDF files in Google Drive.
"""

import io
import re
from typing import Optional

from src.utils.config import settings
from src.utils.logger import logger

# Try to import PDF libraries
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2 import service_account
    GOOGLE_API_AVAILABLE = True
except ImportError:
    GOOGLE_API_AVAILABLE = False


class PDFProcessor:
    """
    Processor for extracting text from PDF files.
    Supports Google Drive URLs and direct file IDs.
    """

    # Patterns for extracting Google Drive file IDs
    DRIVE_PATTERNS = [
        r'/file/d/([a-zA-Z0-9_-]+)',     # Standard Drive URL
        r'id=([a-zA-Z0-9_-]+)',           # URL with id parameter
        r'^([a-zA-Z0-9_-]{25,})$'          # Direct file ID
    ]

    def __init__(self):
        self.drive_service = None
        self._init_drive_service()

    def _init_drive_service(self):
        """Initialize Google Drive service."""
        if not GOOGLE_API_AVAILABLE:
            logger.warning("google_api_not_available",
                          message="Install google-api-python-client for Drive support")
            return

        try:
            credentials = service_account.Credentials.from_service_account_file(
                settings.google_credentials_file,
                scopes=['https://www.googleapis.com/auth/drive.readonly']
            )
            self.drive_service = build('drive', 'v3', credentials=credentials)
            logger.info("drive_service_initialized")
        except Exception as e:
            logger.error("drive_service_init_failed", error=str(e))

    def extract_file_id(self, url_or_id: str) -> Optional[str]:
        """
        Extract Google Drive file ID from URL or return direct ID.

        Args:
            url_or_id: Google Drive URL or file ID

        Returns:
            File ID or None if not found
        """
        if not url_or_id:
            return None

        text = url_or_id.strip()

        for pattern in self.DRIVE_PATTERNS:
            match = re.search(pattern, text)
            if match:
                return match.group(1)

        return None

    def find_file_by_name(self, filename: str, folder_id: Optional[str] = None) -> Optional[str]:
        """
        Find file in Google Drive by name.

        Args:
            filename: File name to search (can be partial)
            folder_id: Optional parent folder ID to limit search

        Returns:
            File ID if found, None otherwise
        """
        if not self.drive_service:
            raise RuntimeError("Google Drive service not initialized")

        # Clean up filename
        clean_name = filename.strip()
        # Escape single quotes for query
        escaped_name = clean_name.replace("'", "\\'")

        # Build query - search for PDF files with matching name
        query = f"name contains '{escaped_name}' and mimeType = 'application/pdf' and trashed = false"
        if folder_id:
            query += f" and '{folder_id}' in parents"

        try:
            results = self.drive_service.files().list(
                q=query,
                fields='files(id, name)',
                pageSize=10,
                orderBy='modifiedTime desc'
            ).execute()

            files = results.get('files', [])
            if files:
                # Return first match (most recently modified)
                logger.info("file_found_by_name", 
                           search_name=filename,
                           found_name=files[0]['name'], 
                           file_id=files[0]['id'])
                return files[0]['id']

            logger.warning("file_not_found_by_name", filename=filename)
            return None

        except Exception as e:
            logger.error("find_file_by_name_failed", filename=filename, error=str(e))
            raise

    def resolve_file_id(self, url_or_id_or_name: str) -> str:
        """
        Resolve file ID from URL, direct ID, or filename.
        Tries in order: URL patterns, direct ID, then name search.

        Args:
            url_or_id_or_name: Google Drive URL, file ID, or filename

        Returns:
            File ID

        Raises:
            ValueError: If file cannot be found
        """
        if not url_or_id_or_name:
            raise ValueError("Empty input: no URL, ID, or filename provided")

        text = url_or_id_or_name.strip()

        # 1. Try to extract from URL patterns
        file_id = self.extract_file_id(text)
        if file_id:
            logger.info("file_id_from_url", file_id=file_id)
            return file_id

        # 2. If it looks like a filename (has extension), search by name
        if '.' in text and not '/' in text and len(text) < 256:
            logger.info("searching_file_by_name", filename=text)
            file_id = self.find_file_by_name(text)
            if file_id:
                return file_id

        raise ValueError(f"Could not resolve file ID from: {url_or_id_or_name}")

    def download_pdf(self, file_id: str) -> bytes:
        """
        Download PDF file from Google Drive.

        Args:
            file_id: Google Drive file ID

        Returns:
            PDF file bytes
        """
        if not self.drive_service:
            raise RuntimeError("Google Drive service not initialized")

        try:
            # Get file metadata
            file_meta = self.drive_service.files().get(
                fileId=file_id,
                fields='name,mimeType'
            ).execute()

            logger.info("downloading_pdf",
                       name=file_meta.get('name'),
                       mime_type=file_meta.get('mimeType'))

            # Download file
            request = self.drive_service.files().get_media(fileId=file_id)
            buffer = io.BytesIO()
            downloader = MediaIoBaseDownload(buffer, request)

            done = False
            while not done:
                status, done = downloader.next_chunk()

            buffer.seek(0)
            return buffer.read()

        except Exception as e:
            logger.error("pdf_download_failed", file_id=file_id, error=str(e))
            raise

    def extract_text_from_bytes(self, pdf_bytes: bytes) -> str:
        """
        Extract text from PDF bytes using PyMuPDF.

        Args:
            pdf_bytes: PDF file content

        Returns:
            Extracted text
        """
        if not PYMUPDF_AVAILABLE:
            raise RuntimeError(
                "PyMuPDF not installed. Install with: pip install pymupdf"
            )

        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            text_parts = []

            for page_num, page in enumerate(doc):
                page_text = page.get_text()
                if page_text.strip():
                    text_parts.append(page_text)
                    logger.debug("pdf_page_extracted",
                               page=page_num + 1,
                               chars=len(page_text))

            doc.close()

            full_text = "\n\n".join(text_parts)
            logger.info("pdf_text_extracted",
                       pages=len(text_parts),
                       total_chars=len(full_text))

            return full_text

        except Exception as e:
            logger.error("pdf_text_extraction_failed", error=str(e))
            raise

    def extract_text(self, url_or_id: str) -> str:
        """
        Extract text from PDF file in Google Drive.

        Args:
            url_or_id: Google Drive URL, file ID, or filename

        Returns:
            Extracted text from PDF
        """
        # Resolve file ID (supports URLs, IDs, and filenames)
        file_id = self.resolve_file_id(url_or_id)

        logger.info("pdf_extraction_start", file_id=file_id)

        # Download PDF
        pdf_bytes = self.download_pdf(file_id)

        # Extract text
        text = self.extract_text_from_bytes(pdf_bytes)

        if not text or len(text) < 10:
            raise ValueError(
                "Extracted text is too short or empty. "
                "The PDF might be a scanned image without OCR."
            )

        return text

    def extract_text_with_ocr(self, url_or_id: str) -> str:
        """
        Extract text from PDF using Google Drive's OCR capabilities.
        Converts PDF to Google Doc and extracts text.

        This is useful for scanned PDFs that PyMuPDF can't handle.

        Args:
            url_or_id: Google Drive URL, file ID, or filename

        Returns:
            Extracted text
        """
        if not self.drive_service:
            raise RuntimeError("Google Drive service not initialized")

        # Resolve file ID (supports URLs, IDs, and filenames)
        file_id = self.resolve_file_id(url_or_id)

        try:
            # Download PDF
            pdf_bytes = self.download_pdf(file_id)

            # Upload as Google Doc with OCR
            from googleapiclient.http import MediaIoBaseUpload

            file_metadata = {
                'name': 'temp_ocr_conversion',
                'mimeType': 'application/vnd.google-apps.document'
            }

            media = MediaIoBaseUpload(
                io.BytesIO(pdf_bytes),
                mimetype='application/pdf',
                resumable=True
            )

            # Create temp doc with OCR
            temp_file = self.drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()

            temp_doc_id = temp_file['id']
            logger.info("ocr_temp_doc_created", doc_id=temp_doc_id)

            try:
                # Export as plain text
                text_content = self.drive_service.files().export(
                    fileId=temp_doc_id,
                    mimeType='text/plain'
                ).execute()

                # Decode if bytes
                if isinstance(text_content, bytes):
                    text = text_content.decode('utf-8')
                else:
                    text = str(text_content)

                logger.info("ocr_text_extracted", chars=len(text))
                return text

            finally:
                # Clean up temp file
                try:
                    self.drive_service.files().delete(fileId=temp_doc_id).execute()
                    logger.debug("ocr_temp_doc_deleted", doc_id=temp_doc_id)
                except Exception as e:
                    logger.warning("ocr_temp_doc_delete_failed", error=str(e))

        except Exception as e:
            logger.error("ocr_extraction_failed", error=str(e))
            raise


# Singleton instance
_pdf_processor: Optional[PDFProcessor] = None


def get_pdf_processor() -> PDFProcessor:
    """Get or create the PDF processor singleton."""
    global _pdf_processor
    if _pdf_processor is None:
        _pdf_processor = PDFProcessor()
    return _pdf_processor
