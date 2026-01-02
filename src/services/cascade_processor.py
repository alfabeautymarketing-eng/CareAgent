"""
Cascade rules processor for certification sheet.

Handles cascading field updates when trigger columns change.
Ported from GAS 03Синхронизация.js _runCertificationCascade.
"""
import re
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass

from src.utils.config import load_project_config
from src.utils.logger import logger


@dataclass
class CascadeChange:
    """Represents a single field change from cascade."""
    column: str
    old_value: str
    new_value: str


@dataclass
class CascadeResult:
    """Result of cascade processing."""
    row: int
    changes: List[CascadeChange]
    applied: bool
    message: str = ""


class CascadeProcessor:
    """
    Processes cascade rules for certification sheet.

    When trigger columns change, recalculates derived fields:
    - VOL_EN: Volume in English (мл→ml, гр→g, etc.)
    - DS_NAME: Combined RUS/ENG name
    - INV_RU: Invoice name in Russian
    - INV_EN: Invoice name in English
    """

    # Default source headers that trigger cascade
    DEFAULT_SOURCE_HEADERS = [
        "Наименования рус по ДС",
        "Наименования англ по ДС",
        "Объём",
        "Код ТН ВЭД",
    ]

    # Target field names
    TARGET_FIELDS = {
        "RUS": "Наименования рус по ДС",
        "ENG": "Наименования англ по ДС",
        "VOL": "Объём",
        "TNVED": "Код ТН ВЭД",
        "VOL_EN": "Объём англ.",
        "DS_NAME": "Наименование ДС",
        "INV_RU": "Наименование для инвойса",
        "INV_EN": "Наименование для инвойса Англ",
    }

    # Default volume replacements (Russian → English)
    DEFAULT_VOLUME_REPLACEMENTS = [
        {"from": "мл", "to": "ml"},
        {"from": "гр", "to": "g"},
        {"from": "шт", "to": "pcs"},
        {"from": "мг", "to": "mg"},
        {"from": "Тестер", "to": "Tester"},
        {"from": "шт. х", "to": "*"},
        {"from": "шт.х", "to": "*"},
    ]

    def __init__(self, sheets_service=None):
        """
        Initialize cascade processor.

        Args:
            sheets_service: Google Sheets service for reading/writing
        """
        self.sheets = sheets_service
        self._load_config()

    def _load_config(self):
        """Load cascade rules from config."""
        try:
            # Try to load from cascade_rules.yaml
            import yaml
            from pathlib import Path

            config_path = Path("config/rules/cascade_rules.yaml")
            if config_path.exists():
                with open(config_path, "r", encoding="utf-8") as f:
                    config = yaml.safe_load(f)

                cert_config = config.get("certification_cascade", {})
                self.source_headers = cert_config.get(
                    "source_headers", self.DEFAULT_SOURCE_HEADERS
                )

                # Load volume replacements
                transforms = cert_config.get("transformations", {})
                vol_config = transforms.get("volume_en", {})
                self.volume_replacements = vol_config.get(
                    "replacements", self.DEFAULT_VOLUME_REPLACEMENTS
                )
            else:
                self.source_headers = self.DEFAULT_SOURCE_HEADERS
                self.volume_replacements = self.DEFAULT_VOLUME_REPLACEMENTS

        except Exception as e:
            logger.warning(f"Failed to load cascade config: {e}, using defaults")
            self.source_headers = self.DEFAULT_SOURCE_HEADERS
            self.volume_replacements = self.DEFAULT_VOLUME_REPLACEMENTS

    @staticmethod
    def normalize(text: str) -> str:
        """
        Normalize text for comparison.

        - Lowercase
        - Replace ё with е
        - Strip whitespace
        """
        return str(text or "").strip().lower().replace("ё", "е")

    def is_trigger_header(self, header: str) -> bool:
        """Check if header is a cascade trigger."""
        normalized = self.normalize(header)
        return any(
            self.normalize(h) == normalized
            for h in self.source_headers
        )

    def convert_volume_to_english(self, volume: str) -> str:
        """
        Convert volume string from Russian to English.

        Examples:
            "100 мл" → "100 ml"
            "50 гр" → "50 g"
        """
        result = str(volume or "")

        for replacement in self.volume_replacements:
            try:
                pattern = replacement.get("from", "")
                to = replacement.get("to", "")
                if pattern:
                    result = re.sub(pattern, to, result, flags=re.IGNORECASE)
            except Exception as e:
                logger.warning(f"Volume replacement error: {e}")

        # Normalize whitespace
        result = re.sub(r"\s+", " ", result).strip()
        return result

    def build_ds_name(self, rus_name: str, eng_name: str) -> str:
        """
        Build DS name from Russian and English names.

        Rules:
        - If Russian name ends with comma → "{rus} {eng}"
        - Otherwise → "{rus} / {eng}"
        - If only one is present, use that one
        """
        rus = str(rus_name or "").strip()
        eng = str(eng_name or "").strip()

        if rus and eng:
            # Check if Russian name ends with comma
            if re.search(r",\s*$", rus):
                return f"{rus} {eng}"
            else:
                return f"{rus} / {eng}"
        else:
            return rus or eng

    def build_invoice_name_ru(
        self, ds_name: str, volume: str, tn_ved: str
    ) -> str:
        """
        Build Russian invoice name.

        Format: "{ds_name} {volume}\nКод ТН ВЭД: {tn_ved}"
        """
        result = f"{ds_name} {volume}".strip() if ds_name else str(volume or "").strip()

        if tn_ved:
            result += f"\nКод ТН ВЭД: {tn_ved}"

        return result

    def build_invoice_name_en(
        self, eng_name: str, volume_en: str, tn_ved: str
    ) -> str:
        """
        Build English invoice name.

        Format: "{eng_name} {volume_en}\nCode: {tn_ved}"
        """
        if eng_name:
            result = f"{eng_name} {volume_en}".strip()
        else:
            result = str(volume_en or "").strip()

        if tn_ved:
            result += f"\nCode: {tn_ved}"

        return result

    def process_row(
        self,
        row_data: Dict[str, Any],
        changed_column: str,
    ) -> CascadeResult:
        """
        Process cascade for a single row.

        Args:
            row_data: Dictionary of column -> value for the row
            changed_column: Name of the column that triggered the cascade

        Returns:
            CascadeResult with list of changes
        """
        changes = []

        # Check if changed column is a trigger
        if not self.is_trigger_header(changed_column):
            return CascadeResult(
                row=0,
                changes=[],
                applied=False,
                message=f"Column '{changed_column}' is not a cascade trigger"
            )

        # Extract source values
        def get_value(field_key: str) -> str:
            field_name = self.TARGET_FIELDS.get(field_key, "")
            # Try exact match first
            if field_name in row_data:
                return str(row_data[field_name] or "")
            # Try normalized match
            normalized = self.normalize(field_name)
            for key, value in row_data.items():
                if self.normalize(key) == normalized:
                    return str(value or "")
            return ""

        rus_name = get_value("RUS")
        eng_name = get_value("ENG")
        volume = get_value("VOL")
        tn_ved = get_value("TNVED")
        current_vol_en = get_value("VOL_EN")

        # Build new values
        new_ds_name = self.build_ds_name(rus_name, eng_name)

        # Volume EN - only recalculate if volume column changed
        normalized_changed = self.normalize(changed_column)
        normalized_vol = self.normalize(self.TARGET_FIELDS["VOL"])

        if normalized_changed == normalized_vol:
            new_vol_en = self.convert_volume_to_english(volume)
        else:
            new_vol_en = current_vol_en

        # Invoice names
        new_inv_ru = self.build_invoice_name_ru(new_ds_name, volume, tn_ved)
        new_inv_en = self.build_invoice_name_en(eng_name, new_vol_en, tn_ved)

        # Collect changes (only if different)
        def add_change(field_key: str, new_value: str):
            field_name = self.TARGET_FIELDS.get(field_key, "")
            old_value = get_value(field_key)
            if self.normalize(old_value) != self.normalize(new_value):
                changes.append(CascadeChange(
                    column=field_name,
                    old_value=old_value,
                    new_value=new_value
                ))

        # DS Name
        if rus_name or eng_name:
            add_change("DS_NAME", new_ds_name)

        # Volume EN (only if volume changed)
        if normalized_changed == normalized_vol:
            add_change("VOL_EN", new_vol_en)

        # Invoice names
        add_change("INV_RU", new_inv_ru)
        add_change("INV_EN", new_inv_en)

        return CascadeResult(
            row=0,
            changes=changes,
            applied=len(changes) > 0,
            message=f"Processed {len(changes)} changes"
        )

    async def process_single_row(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        row: int,
        changed_column: str,
        new_value: Any = None,
        dry_run: bool = False
    ) -> CascadeResult:
        """
        Process cascade for a single row in Google Sheets.

        Args:
            spreadsheet_id: ID of the spreadsheet
            sheet_name: Name of the sheet (usually "Сертификация")
            row: Row number (1-indexed)
            changed_column: Column that was changed
            new_value: New value (optional, will read from sheet if not provided)
            dry_run: If True, don't write changes

        Returns:
            CascadeResult with changes
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            # Read row data
            row_data = await self._read_row_data(spreadsheet_id, sheet_name, row)

            # If new_value provided, update the row_data
            if new_value is not None:
                row_data[changed_column] = new_value

            # Process cascade
            result = self.process_row(row_data, changed_column)
            result.row = row

            if result.applied and not dry_run:
                # Write changes to sheet
                await self._write_changes(
                    spreadsheet_id, sheet_name, row, result.changes
                )

            return result

        except Exception as e:
            logger.error(f"Cascade processing error for row {row}: {e}")
            return CascadeResult(
                row=row,
                changes=[],
                applied=False,
                message=f"Error: {str(e)}"
            )

    async def process_all_rows(
        self,
        spreadsheet_id: str,
        sheet_name: str = "Сертификация",
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Process cascade for all rows in a sheet.

        Args:
            spreadsheet_id: ID of the spreadsheet
            sheet_name: Name of the sheet
            dry_run: If True, don't write changes

        Returns:
            Summary of all changes
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            # Read all data
            all_data = await self._read_sheet_data(spreadsheet_id, sheet_name)
            headers = all_data[0] if all_data else []

            total_rows = len(all_data) - 1  # Exclude header
            processed = 0
            changed = 0
            errors = []

            # Use "Объём" as trigger to process all fields
            trigger_column = self.TARGET_FIELDS["VOL"]

            for i, row_values in enumerate(all_data[1:], start=2):
                try:
                    row_data = dict(zip(headers, row_values))
                    result = self.process_row(row_data, trigger_column)
                    result.row = i

                    if result.applied:
                        changed += 1
                        if not dry_run:
                            await self._write_changes(
                                spreadsheet_id, sheet_name, i, result.changes
                            )

                    processed += 1

                except Exception as e:
                    errors.append(f"Row {i}: {str(e)}")
                    logger.warning(f"Error processing row {i}: {e}")

            return {
                "status": "success",
                "total_rows": total_rows,
                "processed": processed,
                "changed": changed,
                "dry_run": dry_run,
                "errors": errors
            }

        except Exception as e:
            logger.error(f"Cascade processing error: {e}")
            return {
                "status": "error",
                "message": str(e)
            }

    async def _read_row_data(
        self, spreadsheet_id: str, sheet_name: str, row: int
    ) -> Dict[str, Any]:
        """Read a single row as dictionary."""
        # Read headers from row 1 and data from specified row
        range_headers = f"'{sheet_name}'!1:1"
        range_data = f"'{sheet_name}'!{row}:{row}"

        headers_result = await self.sheets.get(
            spreadsheet_id=spreadsheet_id,
            range=range_headers,
            value_render_option="FORMATTED_VALUE"
        )
        headers = headers_result.get("values", [[]])[0]

        data_result = await self.sheets.get(
            spreadsheet_id=spreadsheet_id,
            range=range_data,
            value_render_option="FORMATTED_VALUE"
        )
        values = data_result.get("values", [[]])[0]

        # Pad values to match headers length
        while len(values) < len(headers):
            values.append("")

        return dict(zip(headers, values))

    async def _read_sheet_data(
        self, spreadsheet_id: str, sheet_name: str
    ) -> List[List[Any]]:
        """Read all data from sheet."""
        result = await self.sheets.get(
            spreadsheet_id=spreadsheet_id,
            range=f"'{sheet_name}'",
            value_render_option="FORMATTED_VALUE"
        )
        return result.get("values", [])

    async def _write_changes(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        row: int,
        changes: List[CascadeChange]
    ):
        """Write cascade changes to sheet."""
        if not changes:
            return

        # Get headers to find column indices
        headers_result = await self.sheets.get(
            spreadsheet_id=spreadsheet_id,
            range=f"'{sheet_name}'!1:1",
            value_render_option="FORMATTED_VALUE"
        )
        headers = headers_result.get("values", [[]])[0]

        # Build header to column index map (normalized)
        header_map = {}
        for i, h in enumerate(headers):
            normalized = self.normalize(h)
            if normalized:
                header_map[normalized] = i + 1  # 1-indexed

        # Prepare batch update
        updates = []
        for change in changes:
            col_idx = header_map.get(self.normalize(change.column))
            if col_idx:
                col_letter = self._col_index_to_letter(col_idx)
                cell_range = f"'{sheet_name}'!{col_letter}{row}"
                updates.append({
                    "range": cell_range,
                    "values": [[change.new_value]]
                })

        if updates:
            await self.sheets.batch_update(
                spreadsheet_id=spreadsheet_id,
                data=updates,
                value_input_option="USER_ENTERED"
            )

            logger.info(
                f"Cascade: wrote {len(updates)} changes to row {row}"
            )

    @staticmethod
    def _col_index_to_letter(index: int) -> str:
        """Convert 1-based column index to letter (1=A, 27=AA)."""
        result = ""
        while index > 0:
            index -= 1
            result = chr(65 + index % 26) + result
            index //= 26
        return result


# Singleton instance
_cascade_processor: Optional[CascadeProcessor] = None


def get_cascade_processor(sheets_service=None) -> CascadeProcessor:
    """Get or create CascadeProcessor singleton."""
    global _cascade_processor
    if _cascade_processor is None:
        _cascade_processor = CascadeProcessor(sheets_service)
    return _cascade_processor
