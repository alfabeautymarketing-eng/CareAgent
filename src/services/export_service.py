"""
Export service for promotions and sets.

Exports data from "Заказ" sheet to target documents:
- Promotions: rows where "АКЦИИ" column has value
- Sets: rows where "Набор" column has value

Ported from GAS 11Выгрузка акции и наборы.js
"""
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from datetime import datetime, timedelta
from enum import Enum

from src.utils.logger import logger


class ExportType(str, Enum):
    """Type of export."""
    PROMOTIONS = "promotions"
    SETS = "sets"


@dataclass
class ExportResult:
    """Result of export operation."""
    export_type: str
    exported_rows: int
    target_doc_id: str
    target_sheet_name: str
    target_url: str
    message: str


class ExportService:
    """
    Service for exporting promotions and sets data.

    Reads data from "Заказ" sheet, enriches with data from
    "Главная" and "Сертификация", then writes to target document.
    """

    # Target document IDs for each project
    TARGET_DOCS = {
        "sk": "1YkGP-1Ipn7qLMKJyxLtm3ATrOhCxO2OuFLMt5WK8tsg",
        "ss": "1Q20jk9Cy8gIEJyKQ2-Ph34qqX3Y_oEdKAOK-o_oaFHQ",
        "mt": "140vuIAJ1dcuAoc10T5EnIFjx1lUq7e7oroBJlBs1TDA"
    }

    # Headers for export sheets
    EXPORT_HEADERS = [
        "ID",
        "Статус",
        "Арт. Рус",
        "Арт. произв.",
        "Наименование ДС",
        "Объём",
        "Категория товара",
        "Закупочная цена, ₽",
        "DDP  -МОСКВА,  ₽",
        "Утвержденная цена ОПТ ₽",
        "Утвержденная РРЦ",
        "Цена дистрибьютора -30%,  ₽",
        "Остаток на текущий день",
        "АКЦИИ",  # or "Набор" depending on type
        "условие",
        "коментарий",
        "Срок"
    ]

    # Russian month names
    MONTH_NAMES = [
        "январь", "февраль", "март", "апрель", "май", "июнь",
        "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"
    ]

    def __init__(self, sheets_service=None):
        """
        Initialize export service.

        Args:
            sheets_service: Google Sheets service for reading/writing
        """
        self.sheets = sheets_service

    def get_target_doc_id(self, project: str) -> Optional[str]:
        """Get target document ID for project."""
        return self.TARGET_DOCS.get(project.lower())

    def get_target_sheet_name(self, export_type: ExportType) -> str:
        """
        Generate target sheet name with next month.

        Examples:
            "Акции январь"
            "Наборы февраль"
        """
        # Get next month
        now = datetime.now()
        next_month = now.replace(day=1) + timedelta(days=32)
        next_month = next_month.replace(day=1)

        month_index = next_month.month - 1  # 0-indexed
        month_name = self.MONTH_NAMES[month_index]

        prefix = "Акции " if export_type == ExportType.PROMOTIONS else "Наборы "
        return prefix + month_name

    async def get_order_data(
        self,
        spreadsheet_id: str,
        export_type: ExportType
    ) -> List[Dict[str, Any]]:
        """
        Get data from "Заказ" sheet for export.

        Args:
            spreadsheet_id: Source spreadsheet ID
            export_type: Type of export (promotions or sets)

        Returns:
            List of row dictionaries
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        # Read order sheet
        result = await self.sheets.get(
            spreadsheet_id=spreadsheet_id,
            range="'Заказ'"
        )
        values = result.get("values", [])

        if not values or len(values) < 2:
            return []

        headers = values[0]
        data_rows = values[1:]

        # Determine column names based on export type
        if export_type == ExportType.PROMOTIONS:
            main_column = "АКЦИИ"
            condition_column = "условие#"
            comment_column = "коментарий#"
            term_column = "Срок#"
        else:
            main_column = "Набор"
            condition_column = "условие"
            comment_column = "коментарий"
            term_column = "Срок"

        # Find column indexes
        def find_index(name: str) -> int:
            try:
                return headers.index(name)
            except ValueError:
                # Try partial match
                for i, h in enumerate(headers):
                    if str(h).startswith(name):
                        return i
                return -1

        indices = {
            "id": find_index("ID"),
            "status": find_index("Статус"),
            "rus_art": find_index("Арт. Рус"),
            "prod_art": find_index("Арт. произв."),
            "volume": find_index("Объём"),
            "main_column": find_index(main_column),
            "condition": find_index(condition_column),
            "comment": find_index(comment_column),
            "term": find_index(term_column),
            "stock": find_index("Остаток")
        }

        if indices["main_column"] == -1:
            logger.warning(f"Column '{main_column}' not found")
            return []

        # Filter and extract rows
        rows = []
        for row in data_rows:
            main_value = row[indices["main_column"]] if indices["main_column"] < len(row) else ""

            # Only include rows where main column has value
            if main_value and str(main_value).strip():
                rows.append({
                    "id": row[indices["id"]] if indices["id"] >= 0 and indices["id"] < len(row) else "",
                    "status": row[indices["status"]] if indices["status"] >= 0 and indices["status"] < len(row) else "",
                    "rus_art": row[indices["rus_art"]] if indices["rus_art"] >= 0 and indices["rus_art"] < len(row) else "",
                    "prod_art": row[indices["prod_art"]] if indices["prod_art"] >= 0 and indices["prod_art"] < len(row) else "",
                    "volume": row[indices["volume"]] if indices["volume"] >= 0 and indices["volume"] < len(row) else "",
                    "main_value": main_value,
                    "condition": row[indices["condition"]] if indices["condition"] >= 0 and indices["condition"] < len(row) else "",
                    "comment": row[indices["comment"]] if indices["comment"] >= 0 and indices["comment"] < len(row) else "",
                    "term": row[indices["term"]] if indices["term"] >= 0 and indices["term"] < len(row) else "",
                    "stock": row[indices["stock"]] if indices["stock"] >= 0 and indices["stock"] < len(row) else ""
                })

        return rows

    async def get_enrichment_data(
        self,
        spreadsheet_id: str
    ) -> Dict[str, Dict[str, Any]]:
        """
        Get enrichment data from "Главная" and "Сертификация" sheets.

        Returns:
            Dictionary mapping ID to enrichment data
        """
        enrichment = {}

        # Read "Главная" sheet
        try:
            result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range="'Главная'"
            )
            values = result.get("values", [])

            if values and len(values) > 1:
                headers = values[0]
                id_idx = headers.index("ID") if "ID" in headers else -1
                category_idx = headers.index("Категория товара") if "Категория товара" in headers else -1

                if id_idx >= 0 and category_idx >= 0:
                    for row in values[1:]:
                        row_id = str(row[id_idx] if id_idx < len(row) else "").strip()
                        if row_id:
                            enrichment[row_id] = {
                                "category": row[category_idx] if category_idx < len(row) else ""
                            }
        except Exception as e:
            logger.warning(f"Failed to read Главная: {e}")

        # Read "Сертификация" sheet
        try:
            result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range="'Сертификация'"
            )
            values = result.get("values", [])

            if values and len(values) > 1:
                headers = values[0]
                id_idx = headers.index("ID") if "ID" in headers else -1
                ds_name_idx = headers.index("Наименование ДС") if "Наименование ДС" in headers else -1

                if id_idx >= 0 and ds_name_idx >= 0:
                    for row in values[1:]:
                        row_id = str(row[id_idx] if id_idx < len(row) else "").strip()
                        if row_id:
                            if row_id not in enrichment:
                                enrichment[row_id] = {}
                            enrichment[row_id]["ds_name"] = row[ds_name_idx] if ds_name_idx < len(row) else ""
        except Exception as e:
            logger.warning(f"Failed to read Сертификация: {e}")

        return enrichment

    def enrich_rows(
        self,
        rows: List[Dict[str, Any]],
        enrichment: Dict[str, Dict[str, Any]]
    ) -> List[List[Any]]:
        """
        Enrich order rows with data from other sheets.

        Returns:
            2D array ready for writing to sheet
        """
        enriched = []

        for row in rows:
            row_id = str(row.get("id", "")).strip()
            extra = enrichment.get(row_id, {})

            enriched.append([
                row.get("id", ""),
                row.get("status", ""),
                row.get("rus_art", ""),
                row.get("prod_art", ""),
                extra.get("ds_name", ""),
                row.get("volume", ""),
                extra.get("category", ""),
                "",  # Закупочная цена, ₽
                "",  # DDP  -МОСКВА,  ₽
                "",  # Утвержденная цена ОПТ ₽
                "",  # Утвержденная РРЦ
                "",  # Цена дистрибьютора -30%,  ₽
                row.get("stock", ""),
                row.get("main_value", ""),
                row.get("condition", ""),
                row.get("comment", ""),
                row.get("term", "")
            ])

        return enriched

    async def write_to_target(
        self,
        target_doc_id: str,
        sheet_name: str,
        rows: List[List[Any]]
    ) -> str:
        """
        Write data to target document.

        Returns:
            URL of the target sheet
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        # Check if sheet exists
        try:
            metadata = await self.sheets.get_spreadsheet_metadata(target_doc_id)
            sheets_list = metadata.get("sheets", [])
            sheet_exists = any(
                s.get("properties", {}).get("title") == sheet_name
                for s in sheets_list
            )
        except Exception:
            sheet_exists = False

        if sheet_exists:
            # Clear existing sheet
            await self.sheets.clear(
                spreadsheet_id=target_doc_id,
                range=f"'{sheet_name}'"
            )
        else:
            # Create new sheet
            await self.sheets.batch_update_spreadsheet(
                spreadsheet_id=target_doc_id,
                requests=[{
                    "addSheet": {
                        "properties": {
                            "title": sheet_name
                        }
                    }
                }]
            )

        # Write headers
        await self.sheets.update(
            spreadsheet_id=target_doc_id,
            range=f"'{sheet_name}'!A1",
            values=[self.EXPORT_HEADERS],
            value_input_option="USER_ENTERED"
        )

        # Write data
        if rows:
            await self.sheets.update(
                spreadsheet_id=target_doc_id,
                range=f"'{sheet_name}'!A2",
                values=rows,
                value_input_option="USER_ENTERED"
            )

        # Format headers as bold
        try:
            sheet_id = None
            metadata = await self.sheets.get_spreadsheet_metadata(target_doc_id)
            for s in metadata.get("sheets", []):
                if s.get("properties", {}).get("title") == sheet_name:
                    sheet_id = s.get("properties", {}).get("sheetId")
                    break

            if sheet_id is not None:
                await self.sheets.batch_update_spreadsheet(
                    spreadsheet_id=target_doc_id,
                    requests=[{
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 0,
                                "endRowIndex": 1
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {"bold": True}
                                }
                            },
                            "fields": "userEnteredFormat.textFormat.bold"
                        }
                    }]
                )
        except Exception as e:
            logger.warning(f"Failed to format headers: {e}")

        return f"https://docs.google.com/spreadsheets/d/{target_doc_id}/edit#gid=0"

    async def export(
        self,
        spreadsheet_id: str,
        project: str,
        export_type: ExportType,
        target_doc_id: Optional[str] = None,
        dry_run: bool = False
    ) -> ExportResult:
        """
        Export promotions or sets to target document.

        Args:
            spreadsheet_id: Source spreadsheet ID
            project: Project code (mt, sk, ss)
            export_type: Type of export
            target_doc_id: Optional custom target document ID
            dry_run: If True, don't actually write

        Returns:
            ExportResult with statistics
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            logger.info(f"Starting {export_type.value} export for {project}")

            # Determine target document
            if not target_doc_id:
                target_doc_id = self.get_target_doc_id(project)

            if not target_doc_id:
                raise ValueError(f"No target document configured for project: {project}")

            # Get data from order sheet
            rows = await self.get_order_data(spreadsheet_id, export_type)

            if not rows:
                return ExportResult(
                    export_type=export_type.value,
                    exported_rows=0,
                    target_doc_id=target_doc_id,
                    target_sheet_name="",
                    target_url="",
                    message="No data found for export"
                )

            # Get enrichment data
            enrichment = await self.get_enrichment_data(spreadsheet_id)

            # Enrich rows
            enriched_rows = self.enrich_rows(rows, enrichment)

            # Generate target sheet name
            sheet_name = self.get_target_sheet_name(export_type)

            if dry_run:
                return ExportResult(
                    export_type=export_type.value,
                    exported_rows=len(enriched_rows),
                    target_doc_id=target_doc_id,
                    target_sheet_name=sheet_name,
                    target_url=f"https://docs.google.com/spreadsheets/d/{target_doc_id}",
                    message=f"[DRY RUN] Would export {len(enriched_rows)} rows"
                )

            # Write to target
            target_url = await self.write_to_target(
                target_doc_id=target_doc_id,
                sheet_name=sheet_name,
                rows=enriched_rows
            )

            logger.info(f"Export completed: {len(enriched_rows)} rows")

            return ExportResult(
                export_type=export_type.value,
                exported_rows=len(enriched_rows),
                target_doc_id=target_doc_id,
                target_sheet_name=sheet_name,
                target_url=target_url,
                message=f"Successfully exported {len(enriched_rows)} rows"
            )

        except Exception as e:
            logger.error(f"Export failed: {e}")
            raise


# Singleton instance
_export_service: Optional[ExportService] = None


def get_export_service(sheets_service=None) -> ExportService:
    """Get or create ExportService singleton."""
    global _export_service
    if _export_service is None:
        _export_service = ExportService(sheets_service)
    return _export_service
