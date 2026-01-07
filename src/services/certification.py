from typing import Dict, Any, List, Optional
from datetime import datetime
import re
from src.services.sheets import SheetsService
from src.services.drive import get_drive_service
from src.utils.logger import logger

class CertificationService:
    def __init__(self, sheets_service: SheetsService):
        self.sheets_service = sheets_service
        self.drive_service = get_drive_service()
        self.news_sheet_name = "New sert"
        # Shared 353pp folder ID from config
        self.default_parent_folder_id = "1xajuDu91Epgh5iz8tG1F0xvoTAkgE7Du"

    async def structure_documents_353pp(self, spreadsheet_id: str) -> Dict[str, Any]:
        """
        Structure documents for 353pp application.
        Reads 'New sert' sheet, finds 353pp rows, copies docs to folder.
        """
        results = {
            "processed": 0,
            "errors": 0,
            "folder_url": "",
            "details": []
        }
        
        try:
            # 1. Get Data
            try:
                ws = self.sheets_service.get_worksheet(spreadsheet_id, self.news_sheet_name)
            except Exception:
                return {"success": False, "message": f"Sheet '{self.news_sheet_name}' not found. Please create it first."}

            data = ws.get_all_records()
            if not data:
                return {"success": False, "message": "Sheet is empty"}

            # 2. Prepare Target Folder
            # Name: "Заявка 353пп YYYY-MM-DD"
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
            folder_name = f"Заявка 353пп {timestamp}"
            
            project_prefix = "Common" # Detect project?
            # Try to detect project from first row 'Арт. Рус' or similar?
            # Or just use the shared folder.
            
            target_folder = self.drive_service.create_folder(folder_name, self.default_parent_folder_id)
            results["folder_url"] = target_folder.get('webViewLink')

            # 3. Process Rows
            for i, row in enumerate(data):
                row_num = i + 2
                doc_type = str(row.get("Вид документа", "")).lower()
                
                # Filter for 353pp
                if "353" not in doc_type:
                    continue
                    
                art_rus = str(row.get("Арт. Рус", "NoArt")).replace("/", "_")
                product_name = str(row.get("Наименование ДС", "Product"))[:50].replace("/", "_")
                
                # Create Product Folder
                product_folder_name = f"{art_rus} {product_name}"
                product_folder = self.drive_service.create_folder(product_folder_name, target_folder['id'])
                
                # Copy documents
                files_to_copy = {
                    "INCI": row.get("INCI"),
                    "COA": row.get("COA"),
                    "Этикетка": row.get("Этикетка"),
                    "Скан ДС": row.get("Скан ДС"),
                    "Макет": row.get("Макет")
                }
                
                row_success = True
                
                for file_type, link in files_to_copy.items():
                    if link and "drive.google.com" in str(link):
                        try:
                            file_id = self.drive_service.get_file_id_from_url(link)
                            if file_id:
                                # Copy with name prefix
                                new_name = f"{art_rus}_{file_type}"
                                self.drive_service.copy_file(file_id, product_folder['id'], new_name)
                        except Exception as e:
                            logger.error(f"Failed to copy {file_type} for row {row_num}", error=str(e))
                            results["details"].append(f"Row {row_num}: Failed {file_type} - {str(e)}")
                            # Don't fail entire row, just log
                            
                results["processed"] += 1

            return {
                "success": True,
                "message": f"Processed {results['processed']} items",
                "data": results
            }

        except Exception as e:
            logger.error("structure_documents_353pp_failed", error=str(e))
            raise
