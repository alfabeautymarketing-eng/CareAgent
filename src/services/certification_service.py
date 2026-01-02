"""
Certification service for managing certification sheets and documents.

Provides functionality for:
- Creating "New sert" (новинки) sheet from Certification
- Spirit number calculations
- Protocol and DS layout generation (placeholder)

Ported from GAS certification functions.
"""
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from datetime import datetime
from enum import Enum

from src.utils.logger import logger


class CertificationStatus(str, Enum):
    """Status values for certification process."""
    NEW = "новинка"
    IN_PROGRESS = "в работе"
    READY = "готов"
    EXPIRED = "истек"


@dataclass
class CertificationRow:
    """Represents a certification row."""
    id: str
    rus_art: str
    prod_art: str
    name_rus: str
    name_eng: str
    volume: str
    ds_name: str
    ds_number: str
    ds_valid_until: str
    spirit_number: str
    spirit_valid_until: str
    status: str


@dataclass
class NewsSheetResult:
    """Result of news sheet creation."""
    status: str
    rows_created: int
    sheet_name: str
    message: str


class CertificationService:
    """
    Service for certification-related operations.

    Handles:
    - Creating news sheets for new products
    - Spirit calculations
    - Document generation (placeholder for Drive operations)
    """

    # Headers for "New sert" (новинки) sheet
    NEWS_SHEET_HEADERS = [
        "ID",
        "Арт. Рус",
        "Арт. произв.",
        "Наименование рус по ДС",
        "Наименование англ по ДС",
        "Объём",
        "Наименование ДС",
        "Категория товара",
        "Код ТН ВЭД",
        "Статус",
        "Примечание"
    ]

    # Statuses that indicate "new" products for certification
    NEW_PRODUCT_STATUSES = ["Новинка", "новинка", "NEW", "Новый"]

    def __init__(self, sheets_service=None):
        """
        Initialize certification service.

        Args:
            sheets_service: Google Sheets service for reading/writing
        """
        self.sheets = sheets_service

    async def create_news_sheet(
        self,
        spreadsheet_id: str,
        source_sheet: str = "Сертификация",
        target_sheet: str = "New sert",
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Create "New sert" sheet from Certification sheet.

        Filters rows where status indicates new products pending certification.

        Args:
            spreadsheet_id: ID of the spreadsheet
            source_sheet: Source sheet name (default: Сертификация)
            target_sheet: Target sheet name (default: New sert)
            dry_run: If True, don't write changes

        Returns:
            Dictionary with result details
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            logger.info(f"Creating news sheet from {source_sheet}")

            # Read source sheet
            result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{source_sheet}'"
            )
            values = result.get("values", [])

            if not values or len(values) < 2:
                return {
                    "status": "success",
                    "rows_created": 0,
                    "sheet_name": target_sheet,
                    "message": "No data in source sheet"
                }

            headers = values[0]
            data_rows = values[1:]

            # Find column indices
            def find_idx(name: str) -> int:
                normalized = name.lower().strip()
                for i, h in enumerate(headers):
                    if str(h).lower().strip() == normalized:
                        return i
                    if str(h).lower().strip().startswith(normalized):
                        return i
                return -1

            indices = {
                "id": find_idx("ID"),
                "rus_art": find_idx("Арт. Рус"),
                "prod_art": find_idx("Арт. произв."),
                "name_rus": find_idx("Наименования рус по ДС"),
                "name_eng": find_idx("Наименования англ по ДС"),
                "volume": find_idx("Объём"),
                "ds_name": find_idx("Наименование ДС"),
                "category": find_idx("Категория товара"),
                "tn_ved": find_idx("Код ТН ВЭД"),
                "status": find_idx("Статус"),
                "note": find_idx("Примечание"),
            }

            status_idx = indices.get("status", -1)
            if status_idx == -1:
                # Try alternative status column names
                status_idx = find_idx("статус сертификации")
                if status_idx == -1:
                    status_idx = find_idx("сертификация")

            # Filter rows with "new" status
            news_rows = []
            for row in data_rows:
                status_val = ""
                if status_idx >= 0 and status_idx < len(row):
                    status_val = str(row[status_idx]).strip()

                # Check if status indicates new product
                is_new = any(
                    s.lower() in status_val.lower()
                    for s in self.NEW_PRODUCT_STATUSES
                )

                if is_new:
                    # Extract row data
                    def get_val(key: str) -> str:
                        idx = indices.get(key, -1)
                        if idx >= 0 and idx < len(row):
                            return str(row[idx] if row[idx] else "")
                        return ""

                    news_rows.append([
                        get_val("id"),
                        get_val("rus_art"),
                        get_val("prod_art"),
                        get_val("name_rus"),
                        get_val("name_eng"),
                        get_val("volume"),
                        get_val("ds_name"),
                        get_val("category"),
                        get_val("tn_ved"),
                        status_val,
                        get_val("note"),
                    ])

            if dry_run:
                return {
                    "status": "success",
                    "rows_created": len(news_rows),
                    "sheet_name": target_sheet,
                    "message": f"[DRY RUN] Would create {len(news_rows)} rows",
                    "preview": news_rows[:5] if news_rows else []
                }

            if not news_rows:
                return {
                    "status": "success",
                    "rows_created": 0,
                    "sheet_name": target_sheet,
                    "message": "No new products found for certification"
                }

            # Check if target sheet exists
            try:
                metadata = await self.sheets.get_spreadsheet_metadata(spreadsheet_id)
                sheets_list = metadata.get("sheets", [])
                sheet_exists = any(
                    s.get("properties", {}).get("title") == target_sheet
                    for s in sheets_list
                )
            except Exception:
                sheet_exists = False

            if sheet_exists:
                # Clear existing sheet
                await self.sheets.clear(
                    spreadsheet_id=spreadsheet_id,
                    range=f"'{target_sheet}'"
                )
            else:
                # Create new sheet
                await self.sheets.batch_update_spreadsheet(
                    spreadsheet_id=spreadsheet_id,
                    requests=[{
                        "addSheet": {
                            "properties": {
                                "title": target_sheet
                            }
                        }
                    }]
                )

            # Write headers
            await self.sheets.update(
                spreadsheet_id=spreadsheet_id,
                range=f"'{target_sheet}'!A1",
                values=[self.NEWS_SHEET_HEADERS],
                value_input_option="USER_ENTERED"
            )

            # Write data
            await self.sheets.update(
                spreadsheet_id=spreadsheet_id,
                range=f"'{target_sheet}'!A2",
                values=news_rows,
                value_input_option="USER_ENTERED"
            )

            # Format headers
            try:
                sheet_id = None
                metadata = await self.sheets.get_spreadsheet_metadata(spreadsheet_id)
                for s in metadata.get("sheets", []):
                    if s.get("properties", {}).get("title") == target_sheet:
                        sheet_id = s.get("properties", {}).get("sheetId")
                        break

                if sheet_id is not None:
                    await self.sheets.batch_update_spreadsheet(
                        spreadsheet_id=spreadsheet_id,
                        requests=[{
                            "repeatCell": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": 0,
                                    "endRowIndex": 1
                                },
                                "cell": {
                                    "userEnteredFormat": {
                                        "textFormat": {"bold": True},
                                        "backgroundColor": {
                                            "red": 0.9,
                                            "green": 0.9,
                                            "blue": 0.95
                                        }
                                    }
                                },
                                "fields": "userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor"
                            }
                        }]
                    )
            except Exception as e:
                logger.warning(f"Failed to format headers: {e}")

            logger.info(f"Created news sheet with {len(news_rows)} rows")

            return {
                "status": "success",
                "rows_created": len(news_rows),
                "sheet_name": target_sheet,
                "message": f"Created '{target_sheet}' with {len(news_rows)} new products"
            }

        except Exception as e:
            logger.error(f"Failed to create news sheet: {e}")
            raise

    async def calculate_spirit_numbers(
        self,
        spreadsheet_id: str,
        sheet_name: str = "Сертификация",
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Calculate and assign spirit numbers to certification rows.

        Spirit numbers are assigned based on alcohol content and category.

        Args:
            spreadsheet_id: ID of the spreadsheet
            sheet_name: Sheet name
            dry_run: If True, don't write changes

        Returns:
            Dictionary with calculation results
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            logger.info("Calculating spirit numbers")

            # Read certification sheet
            result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{sheet_name}'"
            )
            values = result.get("values", [])

            if not values or len(values) < 2:
                return {
                    "status": "success",
                    "calculated": 0,
                    "message": "No data in sheet"
                }

            headers = values[0]
            data_rows = values[1:]

            # Find relevant column indices
            def find_idx(name: str) -> int:
                normalized = name.lower().strip()
                for i, h in enumerate(headers):
                    if str(h).lower().strip() == normalized:
                        return i
                    if str(h).lower().strip().startswith(normalized):
                        return i
                return -1

            spirit_idx = find_idx("Спирт")
            spirit_num_idx = find_idx("номер спирт")
            id_idx = find_idx("ID")

            if spirit_idx == -1:
                return {
                    "status": "error",
                    "calculated": 0,
                    "message": "Column 'Спирт' not found"
                }

            # Count rows needing spirit numbers
            needs_calculation = 0
            calculated = 0

            for row in data_rows:
                spirit_val = ""
                if spirit_idx < len(row):
                    spirit_val = str(row[spirit_idx]).strip()

                current_num = ""
                if spirit_num_idx >= 0 and spirit_num_idx < len(row):
                    current_num = str(row[spirit_num_idx]).strip()

                # If has spirit content but no number
                if spirit_val and not current_num:
                    needs_calculation += 1

            if dry_run:
                return {
                    "status": "success",
                    "calculated": 0,
                    "needs_calculation": needs_calculation,
                    "message": f"[DRY RUN] {needs_calculation} rows need spirit numbers"
                }

            # NOTE: Actual spirit number assignment requires complex logic
            # involving alcohol content ranges and regulatory requirements.
            # This is a placeholder for the full implementation.

            return {
                "status": "success",
                "calculated": calculated,
                "needs_calculation": needs_calculation,
                "message": f"Spirit calculation: {needs_calculation} rows need attention"
            }

        except Exception as e:
            logger.error(f"Spirit calculation failed: {e}")
            raise

    async def generate_protocols(
        self,
        spreadsheet_id: str,
        protocol_type: str = "353pp",
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Generate certification protocols.

        NOTE: This is a placeholder. Full implementation requires
        Google Drive/Docs integration for document generation.

        Args:
            spreadsheet_id: ID of the spreadsheet
            protocol_type: Type of protocol (353pp, etc.)
            dry_run: If True, don't generate documents

        Returns:
            Dictionary with generation results
        """
        logger.warning("Protocol generation not fully implemented - requires Drive integration")

        return {
            "status": "not_implemented",
            "message": f"Protocol generation ({protocol_type}) requires Google Drive integration",
            "hint": "This functionality needs Google Docs API for document creation"
        }

    async def generate_ds_layouts(
        self,
        spreadsheet_id: str,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Generate DS (Declaration of Safety) layouts.

        NOTE: This is a placeholder. Full implementation requires
        Google Drive/Docs integration for layout generation.

        Args:
            spreadsheet_id: ID of the spreadsheet
            dry_run: If True, don't generate layouts

        Returns:
            Dictionary with generation results
        """
        logger.warning("DS layout generation not fully implemented - requires Drive integration")

        return {
            "status": "not_implemented",
            "message": "DS layout generation requires Google Drive integration",
            "hint": "This functionality needs Google Docs API for layout creation"
        }

    async def collect_documents(
        self,
        spreadsheet_id: str,
        target_folder_id: Optional[str] = None,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Collect certification documents to a folder.

        NOTE: This is a placeholder. Full implementation requires
        Google Drive integration for file operations.

        Args:
            spreadsheet_id: ID of the spreadsheet
            target_folder_id: Target folder ID in Drive
            dry_run: If True, don't copy files

        Returns:
            Dictionary with collection results
        """
        logger.warning("Document collection not fully implemented - requires Drive integration")

        return {
            "status": "not_implemented",
            "message": "Document collection requires Google Drive integration",
            "hint": "This functionality needs Drive API for file operations"
        }


# Singleton instance
_certification_service: Optional[CertificationService] = None


def get_certification_service(sheets_service=None) -> CertificationService:
    """Get or create CertificationService singleton."""
    global _certification_service
    if _certification_service is None:
        _certification_service = CertificationService(sheets_service)
    return _certification_service
