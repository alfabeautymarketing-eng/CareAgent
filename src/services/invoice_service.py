"""
Invoice processing service.

Handles invoice-related operations:
1. formatOrderSheet - normalize numbers, formatting on "Ордер" sheet
2. createFullInvoice - create "Для инвойса" by joining Order, Certification, Labels
3. collectAndCopyDocuments - collect documents and create descriptions

Ported from GAS 04Инвойс.js
"""
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from datetime import datetime
import re

from src.utils.logger import logger


@dataclass
class FormatResult:
    """Result of format operation."""
    rows_processed: int
    columns_formatted: List[str]
    message: str


@dataclass
class InvoiceResult:
    """Result of invoice creation."""
    rows_created: int
    not_found_count: int
    target_sheet: str
    message: str


class InvoiceService:
    """
    Service for invoice processing.

    Handles formatting of order sheet and creation of invoice sheets.
    """

    # Column names for Order sheet
    ORDER_COLUMNS = {
        "QTY": "Кол-во поставка",
        "PRICE": "Цена ед.",
        "DISCOUNT": "Скидка, %",
        "TOTAL": "Итого цена",
        "EXPIRY": "Годен до",
        "PRODUCER_ART": "Арт. произв.",
    }

    # Column names for Certification sheet
    CERT_COLUMNS = {
        "PRODUCER_ART": "Арт. произв.",
        "RUS_ART": "Арт. Рус",
        "STATUS": "Статус",
        "INVOICE_NAME": "Наименование для инвойса",
        "INVOICE_NAME_ENG": "Наименование для инвойса Англ",
        "DS_UNTIL": "ДС до",
        "SPIRIT_UNTIL": "Спирт До",
        "DS_SCAN": "ДС",
        "SPIRIT_SCAN": "Скан спирт",
        "ID": "ID",
        "DS_NUMBER": "номер ДС",
    }

    # Column names for Labels sheet
    LABEL_COLUMNS = {
        "RUS_ART": "Арт. Рус",
        "PURPOSE": "Назначение Этикетка",
        "APPLICATION": "Применение Этикетка",
        "INGREDIENTS": "Активные ингредиенты без %",
    }

    # Target invoice sheet headers
    INVOICE_HEADERS = [
        "ID",
        "Арт. Рус",
        "Арт. произв.",
        "Наименование",
        "Наименование Англ",
        "кол-во",
        "Цена ед.",
        "Сумма",
        "ДС",
        "ДС до",
        "номер ДС",
        "Спирт",
        "Спирт до",
        "статус",
        "Примечание",
        "Назначение Этикетка",
        "Применение Этикетка",
        "Активные ингредиенты без %",
    ]

    def __init__(self, sheets_service=None):
        """
        Initialize invoice service.

        Args:
            sheets_service: Google Sheets service for reading/writing
        """
        self.sheets = sheets_service

    def normalize_number(self, value: Any, is_percent: bool = False) -> Optional[float]:
        """
        Normalize a value to a number.

        Handles:
        - Already numeric values
        - String values with commas/dots
        - Percentage values
        """
        if value is None or value == "":
            return None

        if isinstance(value, (int, float)):
            n = float(value)
            if is_percent and n > 1:
                n = n / 100
            return n

        s = str(value).strip()

        # Check for percent sign
        had_percent = s.endswith("%")
        s = s.replace("%", "").replace("\u00A0", "").replace(" ", "")

        # Handle comma/dot separators
        if "," in s and "." in s:
            s = s.replace(",", "")  # Comma is thousand separator
        else:
            s = s.replace(",", ".")

        try:
            n = float(s)
            if is_percent:
                if had_percent or n > 1:
                    n = n / 100
            return n
        except ValueError:
            return None

    async def format_order_sheet(
        self,
        spreadsheet_id: str,
        sheet_name: str = "Ордер",
        dry_run: bool = False
    ) -> FormatResult:
        """
        Format the Order sheet.

        - Cleans up supplier-specific data
        - Normalizes numeric columns
        - Centers alignment for key columns

        Args:
            spreadsheet_id: ID of the spreadsheet
            sheet_name: Name of the order sheet
            dry_run: If True, don't actually apply changes

        Returns:
            FormatResult with statistics
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            logger.info(f"Formatting order sheet: {sheet_name}")

            # Read sheet data
            result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{sheet_name}'"
            )
            values = result.get("values", [])

            if not values or len(values) < 2:
                return FormatResult(
                    rows_processed=0,
                    columns_formatted=[],
                    message="Sheet is empty"
                )

            headers = values[0]
            data_rows = values[1:]

            # Find column indexes
            def find_col(name: str) -> int:
                try:
                    return headers.index(name)
                except ValueError:
                    return -1

            col_indexes = {
                key: find_col(name)
                for key, name in self.ORDER_COLUMNS.items()
            }

            # Clean up and normalize data
            formatted_data = []
            columns_formatted = []

            for row in data_rows:
                new_row = list(row)

                # Normalize QTY - remove "UN" suffix
                qty_idx = col_indexes.get("QTY", -1)
                if qty_idx >= 0 and qty_idx < len(new_row):
                    val = str(new_row[qty_idx] or "")
                    val = re.sub(r"\s*UN\s*", "", val, flags=re.IGNORECASE).strip()
                    normalized = self.normalize_number(val)
                    if normalized is not None:
                        new_row[qty_idx] = int(normalized) if normalized == int(normalized) else normalized

                # Normalize PRICE
                price_idx = col_indexes.get("PRICE", -1)
                if price_idx >= 0 and price_idx < len(new_row):
                    normalized = self.normalize_number(new_row[price_idx])
                    if normalized is not None:
                        new_row[price_idx] = normalized

                # Normalize TOTAL
                total_idx = col_indexes.get("TOTAL", -1)
                if total_idx >= 0 and total_idx < len(new_row):
                    normalized = self.normalize_number(new_row[total_idx])
                    if normalized is not None:
                        new_row[total_idx] = normalized

                # Normalize DISCOUNT as percent
                discount_idx = col_indexes.get("DISCOUNT", -1)
                if discount_idx >= 0 and discount_idx < len(new_row):
                    normalized = self.normalize_number(new_row[discount_idx], is_percent=True)
                    if normalized is not None:
                        new_row[discount_idx] = normalized

                formatted_data.append(new_row)

            # Track which columns were formatted
            for key, idx in col_indexes.items():
                if idx >= 0:
                    columns_formatted.append(self.ORDER_COLUMNS[key])

            if not dry_run and formatted_data:
                # Write back formatted data
                await self.sheets.update(
                    spreadsheet_id=spreadsheet_id,
                    range=f"'{sheet_name}'!A2",
                    values=formatted_data,
                    value_input_option="USER_ENTERED"
                )

                # Apply number formats
                # TODO: Apply specific number formats via batchUpdate

            return FormatResult(
                rows_processed=len(formatted_data),
                columns_formatted=columns_formatted,
                message=f"Formatted {len(formatted_data)} rows"
            )

        except Exception as e:
            logger.error(f"Format order sheet failed: {e}")
            raise

    async def get_certification_lookup(
        self,
        spreadsheet_id: str
    ) -> Dict[str, Dict[str, Any]]:
        """Get certification data as lookup by producer article."""
        result = await self.sheets.get(
            spreadsheet_id=spreadsheet_id,
            range="'Сертификация'"
        )
        values = result.get("values", [])

        if not values or len(values) < 2:
            return {}

        headers = values[0]
        data = values[1:]

        # Find column indexes
        def find_col(name: str) -> int:
            try:
                return headers.index(name)
            except ValueError:
                return -1

        col_idx = {key: find_col(name) for key, name in self.CERT_COLUMNS.items()}

        lookup = {}
        for row in data:
            art_idx = col_idx.get("PRODUCER_ART", -1)
            if art_idx < 0 or art_idx >= len(row):
                continue

            key = str(row[art_idx] or "").strip().lower()
            if not key:
                continue

            lookup[key] = {
                "rus_art": row[col_idx["RUS_ART"]] if col_idx["RUS_ART"] >= 0 and col_idx["RUS_ART"] < len(row) else "",
                "status": row[col_idx["STATUS"]] if col_idx["STATUS"] >= 0 and col_idx["STATUS"] < len(row) else "",
                "invoice_name": row[col_idx["INVOICE_NAME"]] if col_idx["INVOICE_NAME"] >= 0 and col_idx["INVOICE_NAME"] < len(row) else "",
                "invoice_name_eng": row[col_idx["INVOICE_NAME_ENG"]] if col_idx["INVOICE_NAME_ENG"] >= 0 and col_idx["INVOICE_NAME_ENG"] < len(row) else "",
                "ds_until": row[col_idx["DS_UNTIL"]] if col_idx["DS_UNTIL"] >= 0 and col_idx["DS_UNTIL"] < len(row) else "",
                "spirit_until": row[col_idx["SPIRIT_UNTIL"]] if col_idx["SPIRIT_UNTIL"] >= 0 and col_idx["SPIRIT_UNTIL"] < len(row) else "",
                "ds_scan": row[col_idx["DS_SCAN"]] if col_idx["DS_SCAN"] >= 0 and col_idx["DS_SCAN"] < len(row) else "",
                "spirit_scan": row[col_idx["SPIRIT_SCAN"]] if col_idx["SPIRIT_SCAN"] >= 0 and col_idx["SPIRIT_SCAN"] < len(row) else "",
                "id": row[col_idx["ID"]] if col_idx["ID"] >= 0 and col_idx["ID"] < len(row) else "",
                "ds_number": row[col_idx["DS_NUMBER"]] if col_idx["DS_NUMBER"] >= 0 and col_idx["DS_NUMBER"] < len(row) else "",
            }

        return lookup

    async def get_labels_lookup(
        self,
        spreadsheet_id: str
    ) -> Dict[str, Dict[str, Any]]:
        """Get labels data as lookup by Russian article."""
        result = await self.sheets.get(
            spreadsheet_id=spreadsheet_id,
            range="'Этикетки'"
        )
        values = result.get("values", [])

        if not values or len(values) < 2:
            return {}

        headers = values[0]
        data = values[1:]

        def find_col(name: str) -> int:
            try:
                return headers.index(name)
            except ValueError:
                return -1

        col_idx = {key: find_col(name) for key, name in self.LABEL_COLUMNS.items()}

        lookup = {}
        for row in data:
            art_idx = col_idx.get("RUS_ART", -1)
            if art_idx < 0 or art_idx >= len(row):
                continue

            key = str(row[art_idx] or "").strip().lower()
            if not key:
                continue

            lookup[key] = {
                "purpose": row[col_idx["PURPOSE"]] if col_idx["PURPOSE"] >= 0 and col_idx["PURPOSE"] < len(row) else "",
                "application": row[col_idx["APPLICATION"]] if col_idx["APPLICATION"] >= 0 and col_idx["APPLICATION"] < len(row) else "",
                "ingredients": row[col_idx["INGREDIENTS"]] if col_idx["INGREDIENTS"] >= 0 and col_idx["INGREDIENTS"] < len(row) else "",
            }

        return lookup

    async def create_full_invoice(
        self,
        spreadsheet_id: str,
        order_sheet: str = "Ордер",
        target_sheet: str = "Для инвойса",
        dry_run: bool = False
    ) -> InvoiceResult:
        """
        Create full invoice sheet.

        Joins data from Order, Certification, and Labels sheets.

        Args:
            spreadsheet_id: ID of the spreadsheet
            order_sheet: Name of the order source sheet
            target_sheet: Name of the target invoice sheet
            dry_run: If True, don't actually create sheet

        Returns:
            InvoiceResult with statistics
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            logger.info(f"Creating full invoice from {order_sheet}")

            # Read order data
            order_result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{order_sheet}'"
            )
            order_values = order_result.get("values", [])

            if not order_values or len(order_values) < 2:
                return InvoiceResult(
                    rows_created=0,
                    not_found_count=0,
                    target_sheet=target_sheet,
                    message="Order sheet is empty"
                )

            order_headers = order_values[0]
            order_data = order_values[1:]

            # Get lookups
            cert_lookup = await self.get_certification_lookup(spreadsheet_id)
            labels_lookup = await self.get_labels_lookup(spreadsheet_id)

            # Find order column indexes
            def find_col(name: str) -> int:
                try:
                    return order_headers.index(name)
                except ValueError:
                    return -1

            order_art_idx = find_col(self.ORDER_COLUMNS["PRODUCER_ART"])
            order_qty_idx = find_col(self.ORDER_COLUMNS["QTY"])
            order_price_idx = find_col(self.ORDER_COLUMNS["PRICE"])
            order_total_idx = find_col(self.ORDER_COLUMNS["TOTAL"])

            # Build invoice data
            invoice_rows = []
            not_found_count = 0

            for row in order_data:
                if order_art_idx < 0 or order_art_idx >= len(row):
                    continue

                prod_art = str(row[order_art_idx] or "").strip()
                if not prod_art:
                    continue

                prod_art_key = prod_art.lower()
                cert = cert_lookup.get(prod_art_key, {})

                if not cert:
                    not_found_count += 1

                # Build invoice row
                inv_row = [""] * len(self.INVOICE_HEADERS)

                # Map to invoice columns
                inv_row[0] = cert.get("id", "")  # ID
                inv_row[1] = cert.get("rus_art", "")  # Арт. Рус
                inv_row[2] = prod_art  # Арт. произв.
                inv_row[3] = cert.get("invoice_name", "") or "!!! НЕ НАЙДЕНО !!!"  # Наименование
                inv_row[4] = cert.get("invoice_name_eng", "")  # Наименование Англ
                inv_row[5] = row[order_qty_idx] if order_qty_idx >= 0 and order_qty_idx < len(row) else ""  # кол-во
                inv_row[6] = row[order_price_idx] if order_price_idx >= 0 and order_price_idx < len(row) else ""  # Цена ед.
                inv_row[7] = row[order_total_idx] if order_total_idx >= 0 and order_total_idx < len(row) else ""  # Сумма
                inv_row[8] = cert.get("ds_scan", "")  # ДС
                inv_row[9] = cert.get("ds_until", "")  # ДС до
                inv_row[10] = cert.get("ds_number", "")  # номер ДС
                inv_row[11] = cert.get("spirit_scan", "")  # Спирт
                inv_row[12] = cert.get("spirit_until", "")  # Спирт до
                inv_row[13] = cert.get("status", "")  # статус
                inv_row[14] = ""  # Примечание

                # Labels lookup by Russian article
                rus_art_key = str(cert.get("rus_art", "")).strip().lower()
                if rus_art_key:
                    labels = labels_lookup.get(rus_art_key, {})
                    inv_row[15] = labels.get("purpose", "")  # Назначение Этикетка
                    inv_row[16] = labels.get("application", "")  # Применение Этикетка
                    inv_row[17] = labels.get("ingredients", "")  # Активные ингредиенты без %

                invoice_rows.append(inv_row)

            if dry_run:
                return InvoiceResult(
                    rows_created=len(invoice_rows),
                    not_found_count=not_found_count,
                    target_sheet=target_sheet,
                    message=f"[DRY RUN] Would create {len(invoice_rows)} rows"
                )

            # Create/clear target sheet
            try:
                await self.sheets.clear(
                    spreadsheet_id=spreadsheet_id,
                    range=f"'{target_sheet}'"
                )
            except Exception:
                # Sheet might not exist, try to create it
                await self.sheets.batch_update_spreadsheet(
                    spreadsheet_id=spreadsheet_id,
                    requests=[{
                        "addSheet": {"properties": {"title": target_sheet}}
                    }]
                )

            # Write headers
            await self.sheets.update(
                spreadsheet_id=spreadsheet_id,
                range=f"'{target_sheet}'!A1",
                values=[self.INVOICE_HEADERS],
                value_input_option="USER_ENTERED"
            )

            # Write data
            if invoice_rows:
                await self.sheets.update(
                    spreadsheet_id=spreadsheet_id,
                    range=f"'{target_sheet}'!A2",
                    values=invoice_rows,
                    value_input_option="USER_ENTERED"
                )

            return InvoiceResult(
                rows_created=len(invoice_rows),
                not_found_count=not_found_count,
                target_sheet=target_sheet,
                message=f"Created {len(invoice_rows)} rows, {not_found_count} not found in certification"
            )

        except Exception as e:
            logger.error(f"Create full invoice failed: {e}")
            raise


# Singleton instance
_invoice_service: Optional[InvoiceService] = None


def get_invoice_service(sheets_service=None) -> InvoiceService:
    """Get or create InvoiceService singleton."""
    global _invoice_service
    if _invoice_service is None:
        _invoice_service = InvoiceService(sheets_service)
    return _invoice_service
