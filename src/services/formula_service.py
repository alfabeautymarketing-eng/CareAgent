"""
Formula service for managing price dynamics and calculation formulas.

Provides functionality for:
- Recalculating price dynamics formulas (EXW ALFASPA, Purchase price, DDP, Growth)
- Updating price calculation formulas (INDEX/MATCH from Price Dynamics)
- Adding new year columns to price dynamics sheet

Ported from GAS 02AutoPrice.js functions.
"""
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from datetime import datetime
import re

from src.utils.logger import logger


@dataclass
class YearBlock:
    """Represents a year block in price dynamics sheet."""
    year: int
    exw_col: int  # Column index (0-based)
    discount_col: int
    alfaspa_col: int
    purchase_col: int
    ddp_col: int
    growth_exw_col: int
    growth_ddp_col: int


@dataclass
class FormulaResult:
    """Result of formula operation."""
    status: str
    blocks_processed: int
    rows_updated: int
    message: str


# Block colors for year columns (matches GAS BLOCK_COLORS)
BLOCK_COLORS = [
    "#cc0000",  # Dark red
    "#e06666",  # Red
    "#f6b26b",  # Orange
    "#ffd966",  # Yellow
    "#93c47d",  # Light green
    "#76a5af",  # Teal
    "#6fa8dc",  # Light blue
    "#6d9eeb",  # Blue
    "#8e7cc3",  # Purple
    "#c27ba0",  # Pink
]

# Column templates for new year block
NEW_YEAR_COLUMN_TEMPLATES = [
    {"header": "EXW {year}, €", "format": "0.00"},
    {"header": "СКИДКА ОТ EXW {year}, %", "format": "0%"},
    {"header": "EXW ALFASPA {year}, €", "format": "0.00"},
    {"header": "Закупочная цена {year}, ₽", "format": "#,##0"},
    {"header": "DDP-МОСКВА {year}, ₽", "format": "#,##0"},
    {"header": "Прирост EXW {prev_short}-{curr_short}, %", "format": "0.0%"},
    {"header": "Прирост DDP-МОСКВА {prev_short}-{curr_short}, %", "format": "0.0%"},
]


class FormulaService:
    """
    Service for formula-related operations.

    Handles:
    - Price dynamics formulas (EXW ALFASPA, Purchase, DDP, Growth calculations)
    - Price calculation formulas (INDEX/MATCH lookups)
    - New year column addition
    """

    def __init__(self, sheets_service=None):
        """
        Initialize formula service.

        Args:
            sheets_service: Google Sheets service for reading/writing
        """
        self.sheets = sheets_service

    def _find_column_index(self, headers: List[str], name: str) -> int:
        """Find column index by header name (case-insensitive, trimmed)."""
        normalized = name.lower().strip()
        for i, h in enumerate(headers):
            header_clean = str(h).strip().lower()
            if header_clean == normalized:
                return i
            if header_clean.startswith(normalized):
                return i
        return -1

    def _find_year_blocks(self, headers: List[str]) -> List[YearBlock]:
        """
        Find all year blocks in headers.

        Year blocks start with "EXW {year}, €" pattern.
        """
        blocks = []
        pattern = re.compile(r'^EXW\s+(\d{4}),\s*€$', re.IGNORECASE)

        for c, header in enumerate(headers):
            match = pattern.match(str(header).strip())
            if match:
                year = int(match.group(1))
                # Each block has 7 columns
                if c + 6 < len(headers):
                    blocks.append(YearBlock(
                        year=year,
                        exw_col=c,
                        discount_col=c + 1,
                        alfaspa_col=c + 2,
                        purchase_col=c + 3,
                        ddp_col=c + 4,
                        growth_exw_col=c + 5,
                        growth_ddp_col=c + 6
                    ))

        # Sort by year descending (newest first)
        blocks.sort(key=lambda b: b.year, reverse=True)
        return blocks

    async def recalculate_price_dynamics_formulas(
        self,
        spreadsheet_id: str,
        sheet_name: str = "Динамика цены",
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Recalculate formulas on Price Dynamics sheet.

        Calculates:
        - EXW ALFASPA = EXW * (1 - discount/100)
        - Purchase price = EXW ALFASPA * currency_rate
        - DDP = Purchase * DDP coefficient
        - Growth EXW = current_year / prev_year - 1
        - Growth DDP = current_year_ddp / prev_year_ddp - 1

        Args:
            spreadsheet_id: ID of the spreadsheet
            sheet_name: Sheet name (default: Динамика цены)
            dry_run: If True, don't write changes

        Returns:
            Dictionary with calculation results
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            logger.info(f"Recalculating price dynamics formulas for sheet: {sheet_name}")

            # Read price dynamics sheet
            result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{sheet_name}'"
            )
            values = result.get("values", [])

            if not values or len(values) < 2:
                return {
                    "status": "success",
                    "blocks_processed": 0,
                    "rows_updated": 0,
                    "message": "No data in sheet"
                }

            headers = [str(h).strip() for h in values[0]]
            data_rows = values[1:]

            # Find year blocks
            blocks = self._find_year_blocks(headers)
            if not blocks:
                return {
                    "status": "success",
                    "blocks_processed": 0,
                    "rows_updated": 0,
                    "message": "No year blocks found in headers"
                }

            # Get currency rates from Справочник
            currency_rates = await self._get_currency_rates(spreadsheet_id)
            ddp_coefficient = await self._get_ddp_coefficient(spreadsheet_id)

            # Calculate values for each block
            updates = []
            rows_updated = 0

            for i, row in enumerate(data_rows):
                row_num = i + 2  # 1-indexed, skip header

                for block_idx, block in enumerate(blocks):
                    prev_block = blocks[block_idx + 1] if block_idx + 1 < len(blocks) else None
                    currency_rate = currency_rates.get(block.year, 1.0)

                    # Get source values
                    exw_val = self._get_numeric(row, block.exw_col)
                    discount_val = self._get_numeric(row, block.discount_col)

                    # Calculate EXW ALFASPA
                    if exw_val is not None:
                        if discount_val is not None and discount_val != 0:
                            alfaspa_val = round(exw_val * (1 - discount_val / 100), 1)
                        else:
                            alfaspa_val = exw_val

                        # Calculate Purchase price
                        purchase_val = round(alfaspa_val * currency_rate, 0)

                        # Calculate DDP
                        ddp_val = round(purchase_val * ddp_coefficient, 0) if ddp_coefficient else None

                        updates.append({
                            "range": f"'{sheet_name}'!{self._col_letter(block.alfaspa_col)}{row_num}",
                            "values": [[alfaspa_val]]
                        })
                        updates.append({
                            "range": f"'{sheet_name}'!{self._col_letter(block.purchase_col)}{row_num}",
                            "values": [[purchase_val]]
                        })
                        if ddp_val is not None:
                            updates.append({
                                "range": f"'{sheet_name}'!{self._col_letter(block.ddp_col)}{row_num}",
                                "values": [[ddp_val]]
                            })

                        # Calculate growth if previous block exists
                        if prev_block:
                            prev_exw = self._get_numeric(row, prev_block.exw_col)
                            prev_ddp = self._get_numeric(row, prev_block.ddp_col)

                            if prev_exw and prev_exw != 0:
                                growth_exw = exw_val / prev_exw - 1
                                updates.append({
                                    "range": f"'{sheet_name}'!{self._col_letter(block.growth_exw_col)}{row_num}",
                                    "values": [[growth_exw]]
                                })

                            if prev_ddp and prev_ddp != 0 and ddp_val:
                                growth_ddp = ddp_val / prev_ddp - 1
                                updates.append({
                                    "range": f"'{sheet_name}'!{self._col_letter(block.growth_ddp_col)}{row_num}",
                                    "values": [[growth_ddp]]
                                })

                        rows_updated += 1

            if dry_run:
                return {
                    "status": "success",
                    "blocks_processed": len(blocks),
                    "rows_updated": rows_updated,
                    "message": f"[DRY RUN] Would update {len(updates)} cells across {len(blocks)} year blocks"
                }

            # Batch update all values
            if updates:
                await self.sheets.batch_update_values(
                    spreadsheet_id=spreadsheet_id,
                    data=updates,
                    value_input_option="USER_ENTERED"
                )

            logger.info(f"Updated {len(updates)} cells in {len(blocks)} year blocks")

            return {
                "status": "success",
                "blocks_processed": len(blocks),
                "rows_updated": rows_updated,
                "message": f"Updated {len(updates)} cells across {len(blocks)} year blocks"
            }

        except Exception as e:
            logger.error(f"Failed to recalculate price dynamics formulas: {e}")
            raise

    async def update_price_calculation_formulas(
        self,
        spreadsheet_id: str,
        price_calc_sheet: str = "Расчет цены",
        price_dynamics_sheet: str = "Динамика цены",
        silent: bool = False,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Update INDEX/MATCH formulas on Price Calculation sheet.

        Pulls data from Price Dynamics sheet based on ID matching:
        - EXW previous year
        - EXW current year
        - EXW ALFASPA current year
        - Purchase price
        - DDP

        Args:
            spreadsheet_id: ID of the spreadsheet
            price_calc_sheet: Target sheet name
            price_dynamics_sheet: Source sheet name
            silent: Suppress notifications
            dry_run: If True, don't write changes

        Returns:
            Dictionary with update results
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            logger.info(f"Updating price calculation formulas from {price_dynamics_sheet}")

            current_year = datetime.now().year
            prev_year = current_year - 1

            # Read both sheets in parallel
            calc_result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{price_calc_sheet}'"
            )
            dynamics_result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{price_dynamics_sheet}'"
            )

            calc_values = calc_result.get("values", [])
            dyn_values = dynamics_result.get("values", [])

            if not calc_values or len(calc_values) < 2:
                return {
                    "status": "success",
                    "rows_updated": 0,
                    "message": "No data in price calculation sheet"
                }

            if not dyn_values or len(dyn_values) < 2:
                return {
                    "status": "success",
                    "rows_updated": 0,
                    "message": "No data in price dynamics sheet"
                }

            calc_headers = [str(h).strip() for h in calc_values[0]]
            dyn_headers = [str(h).strip() for h in dyn_values[0]]
            calc_data = calc_values[1:]
            dyn_data = dyn_values[1:]

            # Find column indices in price calculation sheet
            calc_id_col = self._find_column_index(calc_headers, "ID")
            calc_exw_prev_col = self._find_column_index(calc_headers, "EXW предыдущая, €")
            calc_exw_curr_col = self._find_column_index(calc_headers, "EXW текущая, €")
            calc_alfaspa_col = self._find_column_index(calc_headers, "EXW ALFASPA текущая, €")
            calc_purchase_col = self._find_column_index(calc_headers, "Закупочная цена, ₽")
            calc_ddp_col = self._find_column_index(calc_headers, "DDP-МОСКВА, ₽")

            # Find column indices in price dynamics sheet
            dyn_id_col = self._find_column_index(dyn_headers, "ID")
            dyn_exw_prev_col = self._find_column_index(dyn_headers, f"EXW {prev_year}, €")
            dyn_exw_curr_col = self._find_column_index(dyn_headers, f"EXW {current_year}, €")
            dyn_alfaspa_col = self._find_column_index(dyn_headers, f"EXW ALFASPA {current_year}, €")
            dyn_purchase_col = self._find_column_index(dyn_headers, f"Закупочная цена {current_year}, ₽")
            dyn_ddp_col = self._find_column_index(dyn_headers, f"DDP-МОСКВА {current_year}, ₽")

            if calc_id_col == -1:
                return {
                    "status": "error",
                    "rows_updated": 0,
                    "message": "ID column not found in price calculation sheet"
                }

            if dyn_id_col == -1:
                return {
                    "status": "error",
                    "rows_updated": 0,
                    "message": "ID column not found in price dynamics sheet"
                }

            # Build lookup map from dynamics sheet: ID -> row data
            dyn_lookup = {}
            for row in dyn_data:
                if dyn_id_col < len(row):
                    row_id = str(row[dyn_id_col]).strip()
                    if row_id:
                        dyn_lookup[row_id] = row

            # Prepare column mappings
            mappings = [
                (calc_exw_prev_col, dyn_exw_prev_col, "EXW предыдущая"),
                (calc_exw_curr_col, dyn_exw_curr_col, "EXW текущая"),
                (calc_alfaspa_col, dyn_alfaspa_col, "EXW ALFASPA"),
                (calc_purchase_col, dyn_purchase_col, "Закупочная цена"),
                (calc_ddp_col, dyn_ddp_col, "DDP-МОСКВА"),
            ]

            updates = []
            rows_updated = 0

            for i, calc_row in enumerate(calc_data):
                row_num = i + 2  # 1-indexed, skip header

                if calc_id_col >= len(calc_row):
                    continue

                row_id = str(calc_row[calc_id_col]).strip()
                if not row_id:
                    continue

                dyn_row = dyn_lookup.get(row_id)
                if not dyn_row:
                    continue

                row_has_update = False
                for calc_col, dyn_col, name in mappings:
                    if calc_col == -1 or dyn_col == -1:
                        continue

                    if dyn_col < len(dyn_row):
                        value = dyn_row[dyn_col]
                        if value is not None and str(value).strip():
                            updates.append({
                                "range": f"'{price_calc_sheet}'!{self._col_letter(calc_col)}{row_num}",
                                "values": [[value]]
                            })
                            row_has_update = True

                if row_has_update:
                    rows_updated += 1

            if dry_run:
                return {
                    "status": "success",
                    "rows_updated": rows_updated,
                    "message": f"[DRY RUN] Would update {len(updates)} cells for {rows_updated} rows"
                }

            # Batch update
            if updates:
                await self.sheets.batch_update_values(
                    spreadsheet_id=spreadsheet_id,
                    data=updates,
                    value_input_option="USER_ENTERED"
                )

            logger.info(f"Updated {len(updates)} cells for {rows_updated} rows")

            return {
                "status": "success",
                "rows_updated": rows_updated,
                "message": f"Updated {len(updates)} cells for {rows_updated} rows"
            }

        except Exception as e:
            logger.error(f"Failed to update price calculation formulas: {e}")
            raise

    async def add_new_year_columns(
        self,
        spreadsheet_id: str,
        sheet_name: str = "Динамика цены",
        year: Optional[int] = None,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Add new year columns to price dynamics sheet.

        Inserts 7 columns after "Комментарий":
        - EXW {year}, €
        - СКИДКА ОТ EXW {year}, %
        - EXW ALFASPA {year}, €
        - Закупочная цена {year}, ₽
        - DDP-МОСКВА {year}, ₽
        - Прирост EXW {prev_short}-{curr_short}, %
        - Прирост DDP-МОСКВА {prev_short}-{curr_short}, %

        Args:
            spreadsheet_id: ID of the spreadsheet
            sheet_name: Sheet name (default: Динамика цены)
            year: Target year (default: current year)
            dry_run: If True, don't write changes

        Returns:
            Dictionary with addition results
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            target_year = year or datetime.now().year
            prev_year = target_year - 1
            prev_short = str(prev_year)[-2:]
            curr_short = str(target_year)[-2:]

            logger.info(f"Adding new year columns for {target_year}")

            # Read sheet headers
            result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{sheet_name}'!1:1"
            )
            values = result.get("values", [])

            if not values:
                return {
                    "status": "error",
                    "columns_added": 0,
                    "message": "Sheet is empty"
                }

            headers = [str(h).strip() for h in values[0]]

            # Check if year block already exists
            primary_header = f"EXW {target_year}, €"
            if primary_header in headers:
                return {
                    "status": "exists",
                    "columns_added": 0,
                    "message": f"Year block for {target_year} already exists"
                }

            # Find "Комментарий" column
            comment_idx = -1
            for i, h in enumerate(headers):
                if h.lower().strip() == "комментарий":
                    comment_idx = i
                    break

            if comment_idx == -1:
                return {
                    "status": "error",
                    "columns_added": 0,
                    "message": "Column 'Комментарий' not found"
                }

            # Build new headers
            new_headers = []
            for tpl in NEW_YEAR_COLUMN_TEMPLATES:
                header = tpl["header"].format(
                    year=target_year,
                    prev_short=prev_short,
                    curr_short=curr_short
                )
                new_headers.append(header)

            if dry_run:
                return {
                    "status": "success",
                    "columns_added": len(new_headers),
                    "message": f"[DRY RUN] Would add {len(new_headers)} columns after 'Комментарий'",
                    "headers": new_headers
                }

            # Get sheet ID for structural changes
            metadata = await self.sheets.get_spreadsheet_metadata(spreadsheet_id)
            sheet_id = None
            for s in metadata.get("sheets", []):
                if s.get("properties", {}).get("title") == sheet_name:
                    sheet_id = s.get("properties", {}).get("sheetId")
                    break

            if sheet_id is None:
                return {
                    "status": "error",
                    "columns_added": 0,
                    "message": f"Sheet '{sheet_name}' not found"
                }

            # Insert columns after "Комментарий"
            insert_col = comment_idx + 1  # 0-indexed, insert after comment
            await self.sheets.batch_update_spreadsheet(
                spreadsheet_id=spreadsheet_id,
                requests=[{
                    "insertDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "COLUMNS",
                            "startIndex": insert_col,
                            "endIndex": insert_col + len(new_headers)
                        },
                        "inheritFromBefore": False
                    }
                }]
            )

            # Write new headers
            header_range = f"'{sheet_name}'!{self._col_letter(insert_col)}1:{self._col_letter(insert_col + len(new_headers) - 1)}1"
            await self.sheets.update(
                spreadsheet_id=spreadsheet_id,
                range=header_range,
                values=[new_headers],
                value_input_option="USER_ENTERED"
            )

            # Apply header formatting (bold, background color)
            block_color = self._hex_to_rgb(BLOCK_COLORS[0])
            await self.sheets.batch_update_spreadsheet(
                spreadsheet_id=spreadsheet_id,
                requests=[{
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": insert_col,
                            "endColumnIndex": insert_col + len(new_headers)
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {"bold": True},
                                "backgroundColor": block_color
                            }
                        },
                        "fields": "userEnteredFormat.textFormat.bold,userEnteredFormat.backgroundColor"
                    }
                }]
            )

            logger.info(f"Added {len(new_headers)} columns for year {target_year}")

            return {
                "status": "success",
                "columns_added": len(new_headers),
                "year": target_year,
                "message": f"Added {len(new_headers)} columns for year {target_year}"
            }

        except Exception as e:
            logger.error(f"Failed to add new year columns: {e}")
            raise

    async def _get_currency_rates(self, spreadsheet_id: str) -> Dict[int, float]:
        """Get currency rates from Справочник sheet."""
        try:
            result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range="'Справочник'"
            )
            values = result.get("values", [])

            if not values or len(values) < 2:
                return {}

            headers = [str(h).strip() for h in values[0]]
            year_col = self._find_column_index(headers, "Год")
            currency_col = self._find_column_index(headers, "Курс валюты, €")

            if year_col == -1 or currency_col == -1:
                return {}

            rates = {}
            for row in values[1:]:
                if year_col < len(row) and currency_col < len(row):
                    try:
                        year = int(float(str(row[year_col]).strip()))
                        rate = float(str(row[currency_col]).replace(",", ".").strip())
                        rates[year] = rate
                    except (ValueError, TypeError):
                        continue

            return rates
        except Exception as e:
            logger.warning(f"Failed to get currency rates: {e}")
            return {}

    async def _get_ddp_coefficient(self, spreadsheet_id: str) -> Optional[float]:
        """Get DDP coefficient from Справочник sheet (cell D2)."""
        try:
            result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range="'Справочник'!D2"
            )
            values = result.get("values", [])
            if values and values[0]:
                return float(str(values[0][0]).replace(",", ".").strip())
            return None
        except Exception as e:
            logger.warning(f"Failed to get DDP coefficient: {e}")
            return None

    def _get_numeric(self, row: List[Any], col: int) -> Optional[float]:
        """Get numeric value from row at column index."""
        if col < 0 or col >= len(row):
            return None
        val = row[col]
        if val is None or str(val).strip() == "":
            return None
        try:
            return float(str(val).replace(",", ".").replace(" ", "").replace("\xa0", ""))
        except (ValueError, TypeError):
            return None

    def _col_letter(self, index: int) -> str:
        """Convert 0-based column index to letter (A, B, ..., Z, AA, AB, ...)."""
        result = ""
        index += 1  # Convert to 1-based
        while index > 0:
            index -= 1
            result = chr(65 + index % 26) + result
            index //= 26
        return result

    def _hex_to_rgb(self, hex_color: str) -> Dict[str, float]:
        """Convert hex color to RGB dict for Sheets API."""
        hex_color = hex_color.lstrip("#")
        return {
            "red": int(hex_color[0:2], 16) / 255,
            "green": int(hex_color[2:4], 16) / 255,
            "blue": int(hex_color[4:6], 16) / 255
        }


# Singleton instance
_formula_service: Optional[FormulaService] = None


def get_formula_service(sheets_service=None) -> FormulaService:
    """Get or create FormulaService singleton."""
    global _formula_service
    if _formula_service is None:
        _formula_service = FormulaService(sheets_service)
    return _formula_service
