"""
Price Processor Service

Handles parsing and processing of supplier price lists (Б/З поставщик).
Replaces GAS functions: processMtMainPrice, processMtTesterPrice, processMtSamplesPrice,
processSkPriceSheet, processSkPriceProbes, processSsPriceSheet.
"""

import yaml
import re
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field

from src.utils.logger import logger
from src.services.sheets import SheetsService


@dataclass
class ParsedRow:
    """Single parsed row from price list"""
    id_p: str = ""
    article: str = ""
    name_eng: str = ""
    volume: str = ""
    barcode: str = ""
    units_per_pack: str = ""
    price: float = 0.0
    group: str = ""

    def to_list(self) -> List[Any]:
        return [
            self.id_p,
            self.article,
            self.name_eng,
            self.volume,
            self.barcode,
            self.units_per_pack,
            self.price,
            self.group
        ]


@dataclass
class ProcessedData:
    """Result of parsing a price list"""
    headers: List[str]
    rows: List[ParsedRow]
    articles: List[str] = field(default_factory=list)
    groups: List[str] = field(default_factory=list)

    @property
    def unique_groups(self) -> int:
        return len(set(self.groups))


@dataclass
class SyncResult:
    """Result of syncing with main sheet"""
    assigned_idp: Dict[str, str] = field(default_factory=dict)
    created_rows: List[int] = field(default_factory=list)
    updated_rows: List[int] = field(default_factory=list)
    group_changes: List[Dict[str, Any]] = field(default_factory=list)
    barcode_mismatches: List[Dict[str, Any]] = field(default_factory=list)


class PriceProcessor:
    """
    Service for processing supplier price lists.

    Workflow:
    1. Load project config from YAML
    2. Read source data from external document
    3. Parse rows (detect groups, articles)
    4. Assign ID-P values
    5. Sync with "Главная" sheet
    6. Copy prices to related sheets
    7. Apply formulas
    """

    # Output headers for processed data
    OUTPUT_HEADERS = [
        "ID-P",
        "Арт. произв.",
        "Название ENG прайс произв",
        "Объём",
        "BAR CODE",
        "шт./уп.",
        "Цена",
        "Группа",
    ]

    # Column indices in source Excel (from config)
    DEFAULT_EXCEL_COLUMNS = {
        "CODE": 0,
        "CATEGORY": 1,
        "DESCRIPTION": 2,
        "FORMAT": 3,
        "UNITS": 4,
        "PRICE": 5,
        "EAN": 6
    }

    def __init__(self, sheets_service: SheetsService, logging_service=None):
        self.sheets = sheets_service
        self.logging = logging_service
        self._config_cache: Dict[str, Dict] = {}

    def _log(self, spreadsheet_id: str, category: str, action: str, details: str, status: str = "🔄"):
        """Log action to sheet if logging service available"""
        if self.logging:
            try:
                self.logging.add_log(spreadsheet_id, category, action, details, status)
            except Exception:
                pass
        logger.info(f"[{category}] {action}: {details}")

    def load_config(self, project: str) -> Dict[str, Any]:
        """Load project configuration from YAML file"""
        if project in self._config_cache:
            return self._config_cache[project]

        config_path = Path(__file__).parent.parent.parent / "config" / "projects" / f"{project}.yaml"

        if not config_path.exists():
            raise ValueError(f"Config file not found: {config_path}")

        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)

        self._config_cache[project] = config
        return config

    def get_parser_config(self, project: str, mode: str) -> Dict[str, Any]:
        """Get parser configuration for specific mode"""
        config = self.load_config(project)
        parser_config = config.get("parser", {}).get(mode)

        if not parser_config:
            raise ValueError(f"Parser config not found for project={project}, mode={mode}")

        return parser_config

    async def process(
        self,
        project: str,
        mode: str,
        spreadsheet_id: str,
        source_doc_id: Optional[str] = None,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Main processing method.

        Args:
            project: Project code (mt, sk, ss)
            mode: Processing mode (main, tester, samples, probes)
            spreadsheet_id: Target spreadsheet ID
            source_doc_id: Source document ID (optional, uses config if not provided)
            dry_run: If True, returns preview without writing

        Returns:
            Processing result with statistics
        """
        project = project.lower()
        mode = mode.lower()

        self._log(spreadsheet_id, "ПРАЙС", f"Обработка {project.upper()} {mode}", "Старт", "🚀")

        try:
            # 1. Load config
            config = self.load_config(project)
            parser_config = self.get_parser_config(project, mode)

            # 2. Read source data
            source_data = await self._read_source_data(
                spreadsheet_id,
                project,
                mode,
                source_doc_id,
                config
            )

            if not source_data or not source_data.get("values"):
                return {
                    "status": "error",
                    "message": "Источник данных пуст",
                    "processed_rows": 0
                }

            # 3. Parse data based on project/mode
            processed = self._parse_data(source_data["values"], project, mode, parser_config)

            if not processed.rows:
                return {
                    "status": "error",
                    "message": "Не найдено данных для обработки",
                    "processed_rows": 0
                }

            # 4. If dry_run, return preview
            if dry_run:
                preview = [
                    {
                        "article": row.article,
                        "name": row.name_eng,
                        "volume": row.volume,
                        "price": row.price,
                        "group": row.group
                    }
                    for row in processed.rows[:20]
                ]
                return {
                    "status": "preview",
                    "message": f"Найдено {len(processed.rows)} артикулов в {processed.unique_groups} группах",
                    "processed_rows": len(processed.rows),
                    "groups_found": processed.unique_groups,
                    "preview": preview
                }

            # 5. Clear ID-P columns (only for main mode)
            if mode == "main":
                await self._clear_idp_columns(spreadsheet_id, config)
                await self._clear_price_exw_columns(spreadsheet_id, config)

            # 6. Sync with main sheet
            sync_result = await self._sync_with_main(spreadsheet_id, processed, config, project)

            # 7. Apply assigned ID-P to processed data
            self._apply_assigned_idp(processed, sync_result)

            # 8. Fill ID-P on other sheets
            await self._fill_idp_on_sheets(spreadsheet_id, config)

            # 9. Copy prices to related sheets
            price_multiplier = parser_config.get("price_multiplier", 1)
            await self._copy_prices_to_sheets(spreadsheet_id, processed, config, price_multiplier)

            # 10. Recalculate formulas
            await self._recalculate_formulas(spreadsheet_id, config)

            # 11. Update statuses (only for samples mode)
            if mode == "samples" and parser_config.get("update_statuses_after"):
                await self._update_statuses(spreadsheet_id, config)

            self._log(
                spreadsheet_id,
                "ПРАЙС",
                f"Обработка {project.upper()} {mode} завершена",
                f"Строк: {len(processed.rows)}, групп: {processed.unique_groups}",
                "✅"
            )

            return {
                "status": "success",
                "message": f"Обработано {len(processed.rows)} артикулов",
                "processed_rows": len(processed.rows),
                "groups_found": processed.unique_groups,
                "new_articles": len(sync_result.created_rows),
                "updated_articles": len(sync_result.updated_rows),
                "barcode_mismatches": len(sync_result.barcode_mismatches),
                "errors": []
            }

        except Exception as e:
            logger.error(f"Price processing failed: {e}", exc_info=True)
            self._log(spreadsheet_id, "ПРАЙС", "Ошибка обработки", str(e), "❌")
            return {
                "status": "error",
                "message": str(e),
                "processed_rows": 0,
                "errors": [str(e)]
            }

    async def _read_source_data(
        self,
        spreadsheet_id: str,
        project: str,
        mode: str,
        source_doc_id: Optional[str],
        config: Dict
    ) -> Dict[str, Any]:
        """Read source data from external document"""

        # Get source document ID from config if not provided
        if not source_doc_id:
            # Try to get from config source section (new structure)
            source_config = config.get("source", {})
            source_doc_id = source_config.get("doc_id")

            if not source_doc_id:
                # Fallback: use the same spreadsheet
                source_doc_id = spreadsheet_id

        # Get source sheet name based on mode from config
        source_config = config.get("source", {})
        source_sheets = source_config.get("sheets", {})

        # Map mode to sheet name from config
        mode_to_key = {
            "main": "main",
            "tester": "tester",
            "samples": "samples",
            "probes": "probes"
        }
        sheet_key = mode_to_key.get(mode, "main")
        sheet_name = source_sheets.get(sheet_key, "-Б/З поставщик")

        # Override from parser config if available
        parser_config = config.get("parser", {}).get(mode, {})
        if parser_config.get("sheet_name"):
            sheet_name = parser_config["sheet_name"]

        try:
            ws = self.sheets.get_worksheet(source_doc_id, sheet_name)
            values = ws.get_all_values()

            # Try to get backgrounds for color-based parsing (SK)
            backgrounds = None
            if project == "sk":
                try:
                    backgrounds = ws.get_all_backgrounds()
                except Exception:
                    pass

            return {
                "values": values,
                "backgrounds": backgrounds,
                "sheet_name": sheet_name,
                "source_doc_id": source_doc_id
            }

        except Exception as e:
            logger.error(f"Failed to read source data: {e}")
            raise ValueError(f"Не удалось прочитать данные из листа '{sheet_name}': {e}")

    def _parse_data(
        self,
        values: List[List[Any]],
        project: str,
        mode: str,
        parser_config: Dict
    ) -> ProcessedData:
        """Parse source data based on project and mode"""

        if project == "mt":
            return self._parse_mt_data(values, mode, parser_config)
        elif project == "sk":
            return self._parse_sk_data(values, mode, parser_config)
        elif project == "ss":
            return self._parse_ss_data(values, mode, parser_config)
        else:
            raise ValueError(f"Unknown project: {project}")

    def _parse_mt_data(
        self,
        values: List[List[Any]],
        mode: str,
        parser_config: Dict
    ) -> ProcessedData:
        """Parse MT project data (main, tester, samples)"""

        rows: List[ParsedRow] = []
        articles: List[str] = []
        groups: List[str] = []
        current_group = ""

        # Find header row
        header_row_index = self._find_header_row(values)
        if header_row_index == -1:
            raise ValueError("Не найдена строка заголовков с CODE и DESCRIPTION")

        headers = values[header_row_index]

        # Get column indices
        excel_cols = parser_config.get("excel_columns", self.DEFAULT_EXCEL_COLUMNS)
        code_idx = self._find_column_index(headers, "CODE", excel_cols.get("CODE", 0))
        category_idx = self._find_column_index(headers, "CATEGORY", excel_cols.get("CATEGORY", 1))
        desc_idx = self._find_column_index(headers, "DESCRIPTION", excel_cols.get("DESCRIPTION", 2))
        format_idx = self._find_column_index(headers, "FORMAT", excel_cols.get("FORMAT", 3))
        units_idx = self._find_column_index(headers, "UNITS", excel_cols.get("UNITS", 4))
        price_idx = self._find_column_index(headers, "PRICE", excel_cols.get("PRICE", 5))
        ean_idx = self._find_column_index(headers, "EAN", excel_cols.get("EAN", 6))

        if code_idx == -1 or desc_idx == -1:
            raise ValueError("Не найдены обязательные столбцы CODE и DESCRIPTION")

        # Get transformations for mode
        transformations = parser_config.get("transformations", {})
        volume_prefix = transformations.get("volume_prefix", "")
        group_suffix = transformations.get("group_suffix", "")
        price_multiplier = parser_config.get("price_multiplier", 1)

        # Parse rows
        for i in range(header_row_index + 1, len(values)):
            row = values[i]
            code_value = self._as_trimmed_string(self._safe_get(row, code_idx))
            category_value = self._as_trimmed_string(self._safe_get(row, category_idx)) if category_idx != -1 else ""
            desc_value = self._as_trimmed_string(self._safe_get(row, desc_idx))

            # Group detection: CODE exists AND CATEGORY empty
            if code_value and not category_value and not desc_value:
                current_group = code_value
                continue

            # Article detection: CODE and DESCRIPTION exist
            if code_value and desc_value:
                format_value = self._get_value(row, format_idx)
                units_value = self._get_value(row, units_idx)
                price_value = self._parse_price(self._get_value(row, price_idx)) * price_multiplier
                ean_value = self._get_value(row, ean_idx)

                # Apply transformations
                volume = f"{volume_prefix}{format_value}" if volume_prefix and format_value else format_value
                group = f"{current_group}{group_suffix}" if group_suffix else current_group

                parsed_row = ParsedRow(
                    article=code_value,
                    name_eng=desc_value,
                    volume=volume,
                    barcode=str(ean_value) if ean_value else "",
                    units_per_pack=str(units_value) if units_value else "",
                    price=price_value,
                    group=group
                )

                rows.append(parsed_row)
                articles.append(code_value)
                groups.append(group)

        return ProcessedData(
            headers=self.OUTPUT_HEADERS.copy(),
            rows=rows,
            articles=articles,
            groups=groups
        )

    def _parse_sk_data(
        self,
        values: List[List[Any]],
        mode: str,
        parser_config: Dict
    ) -> ProcessedData:
        """Parse SK project data with color-based group detection"""
        # TODO: Implement color-based parsing for SK
        # For now, use same logic as MT
        return self._parse_mt_data(values, mode, parser_config)

    def _parse_ss_data(
        self,
        values: List[List[Any]],
        mode: str,
        parser_config: Dict
    ) -> ProcessedData:
        """Parse SS project data with markers (-ПРОФ, SAMPLES)"""
        # TODO: Implement marker-based parsing for SS
        # For now, use same logic as MT
        return self._parse_mt_data(values, mode, parser_config)

    def _find_header_row(self, values: List[List[Any]], max_rows: int = 5) -> int:
        """Find row index containing CODE and DESCRIPTION headers"""
        for i in range(min(len(values), max_rows)):
            row = values[i]
            has_code = any("CODE" in str(cell).upper() for cell in row)
            has_desc = any("DESCRIPTION" in str(cell).upper() for cell in row)
            if has_code and has_desc:
                return i
        return 0  # Default to first row

    def _find_column_index(self, headers: List[Any], keyword: str, default: int = -1) -> int:
        """Find column index by keyword in headers"""
        keyword_upper = keyword.upper()
        for i, header in enumerate(headers):
            if keyword_upper in str(header or "").strip().upper():
                return i
        return default if default >= 0 else -1

    def _as_trimmed_string(self, value: Any) -> str:
        """Convert value to trimmed string"""
        if value is None:
            return ""
        return str(value).strip()

    def _safe_get(self, row: List[Any], idx: int, default: Any = "") -> Any:
        """Safely get value from row by index"""
        if idx < 0 or idx >= len(row):
            return default
        return row[idx] if row[idx] is not None else default

    def _get_value(self, row: List[Any], idx: int) -> str:
        """Get string value from row by index"""
        return self._as_trimmed_string(self._safe_get(row, idx))

    def _parse_price(self, value: Any) -> float:
        """Parse price value to float"""
        if value is None or value == "":
            return 0.0
        try:
            # Handle comma as decimal separator
            str_val = str(value).replace(",", ".").replace(" ", "").replace("\xa0", "")
            # Remove currency symbols
            str_val = re.sub(r"[€$₽£]", "", str_val)
            return float(str_val)
        except (ValueError, TypeError):
            return 0.0

    async def _clear_idp_columns(self, spreadsheet_id: str, config: Dict):
        """Clear ID-P column on all target sheets"""
        sheets_to_clear = config.get("base_sheets_for_creation", [])
        sheets_to_clear.append(config.get("sheets", {}).get("primary", "Главная"))

        self._log(spreadsheet_id, "ПРАЙС", "Очистка ID-P", f"Листы: {', '.join(sheets_to_clear)}", "🔄")

        for sheet_name in sheets_to_clear:
            try:
                ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)
                headers = ws.row_values(1)

                idp_col = None
                for i, h in enumerate(headers):
                    if "ID-P" in str(h).upper():
                        idp_col = i + 1
                        break

                if idp_col:
                    num_rows = ws.row_count
                    if num_rows > 1:
                        # Clear ID-P column (from row 2 to end)
                        ws.batch_clear([f"{chr(64 + idp_col)}2:{chr(64 + idp_col)}{num_rows}"])

            except Exception as e:
                logger.warning(f"Failed to clear ID-P on {sheet_name}: {e}")

    async def _clear_price_exw_columns(self, spreadsheet_id: str, config: Dict):
        """Clear ЦЕНА EXW из Б/З column on price dynamics and calculation sheets"""
        sheets_to_clear = [
            config.get("sheets", {}).get("price_dynamics", "Динамика цены"),
            config.get("sheets", {}).get("price_calculation", "Расчет цены")
        ]

        for sheet_name in sheets_to_clear:
            try:
                ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)
                headers = ws.row_values(1)

                price_col = None
                for i, h in enumerate(headers):
                    if "ЦЕНА EXW" in str(h).upper() and "Б/З" in str(h).upper():
                        price_col = i + 1
                        break

                if price_col:
                    num_rows = ws.row_count
                    if num_rows > 1:
                        ws.batch_clear([f"{chr(64 + price_col)}2:{chr(64 + price_col)}{num_rows}"])

            except Exception as e:
                logger.warning(f"Failed to clear price on {sheet_name}: {e}")

    async def _sync_with_main(
        self,
        spreadsheet_id: str,
        processed: ProcessedData,
        config: Dict,
        project: str
    ) -> SyncResult:
        """Sync processed data with main sheet (Главная)"""

        result = SyncResult()
        primary_sheet = config.get("sheets", {}).get("primary", "Главная")
        prefix = config.get("project", {}).get("code", project).upper() + "-"

        try:
            ws = self.sheets.get_worksheet(spreadsheet_id, primary_sheet)
            all_values = ws.get_all_values()

            if not all_values:
                return result

            headers = all_values[0]

            # Find column indices
            id_col = self._find_column_index(headers, "ID", 0)
            idp_col = self._find_column_index(headers, "ID-P", 1)
            article_col = self._find_column_index(headers, "Арт. произв", 5)
            name_col = self._find_column_index(headers, "Название ENG", 8)
            volume_col = self._find_column_index(headers, "Объём", 11)
            barcode_col = self._find_column_index(headers, "BAR CODE", 17)
            units_col = self._find_column_index(headers, "шт./уп", 13)
            price_col = self._find_column_index(headers, "ЦЕНА EXW", -1)
            group_col = self._find_column_index(headers, "Группа", 16)

            # Build article -> row index map
            article_map: Dict[str, int] = {}
            for i, row in enumerate(all_values[1:], start=2):
                art = self._as_trimmed_string(self._safe_get(row, article_col))
                if art:
                    article_map[art] = i

            # Find max ID-P number
            max_idp = 0
            for row in all_values[1:]:
                idp_val = self._as_trimmed_string(self._safe_get(row, idp_col))
                if idp_val and idp_val.startswith(prefix):
                    try:
                        num = int(idp_val[len(prefix):])
                        max_idp = max(max_idp, num)
                    except ValueError:
                        pass

            # Process each parsed row
            updates = []
            new_rows = []

            for parsed_row in processed.rows:
                article = parsed_row.article

                if article in article_map:
                    # Update existing row
                    row_idx = article_map[article]
                    existing_row = all_values[row_idx - 1]

                    # Check for barcode mismatch
                    existing_barcode = self._as_trimmed_string(self._safe_get(existing_row, barcode_col))
                    if existing_barcode and parsed_row.barcode and existing_barcode != parsed_row.barcode:
                        result.barcode_mismatches.append({
                            "article": article,
                            "existing": existing_barcode,
                            "newValue": parsed_row.barcode
                        })

                    # Assign ID-P if empty
                    existing_idp = self._as_trimmed_string(self._safe_get(existing_row, idp_col))
                    if not existing_idp:
                        max_idp += 1
                        new_idp = f"{prefix}{max_idp}"
                        result.assigned_idp[article] = new_idp
                        updates.append({
                            "range": f"{chr(65 + idp_col)}{row_idx}",
                            "values": [[new_idp]]
                        })
                    else:
                        result.assigned_idp[article] = existing_idp

                    # Update price if column exists
                    if price_col >= 0 and parsed_row.price > 0:
                        updates.append({
                            "range": f"{chr(65 + price_col)}{row_idx}",
                            "values": [[parsed_row.price]]
                        })

                    result.updated_rows.append(row_idx)
                else:
                    # Create new row
                    max_idp += 1
                    new_idp = f"{prefix}{max_idp}"
                    result.assigned_idp[article] = new_idp

                    # Build new row data
                    new_row = [""] * len(headers)
                    new_row[id_col] = ""  # ID will be auto-generated
                    new_row[idp_col] = new_idp
                    if article_col >= 0:
                        new_row[article_col] = article
                    if name_col >= 0:
                        new_row[name_col] = parsed_row.name_eng
                    if volume_col >= 0:
                        new_row[volume_col] = parsed_row.volume
                    if barcode_col >= 0:
                        new_row[barcode_col] = parsed_row.barcode
                    if units_col >= 0:
                        new_row[units_col] = parsed_row.units_per_pack
                    if price_col >= 0:
                        new_row[price_col] = parsed_row.price
                    if group_col >= 0:
                        new_row[group_col] = parsed_row.group

                    new_rows.append(new_row)
                    result.created_rows.append(len(all_values) + len(new_rows))

            # Apply updates
            if updates:
                ws.batch_update(updates, value_input_option="USER_ENTERED")

            # Append new rows
            if new_rows:
                ws.append_rows(new_rows, value_input_option="USER_ENTERED")

            logger.info(
                f"Sync complete: {len(result.updated_rows)} updated, {len(result.created_rows)} created"
            )

        except Exception as e:
            logger.error(f"Sync with main failed: {e}", exc_info=True)
            raise

        return result

    def _apply_assigned_idp(self, processed: ProcessedData, sync_result: SyncResult):
        """Apply assigned ID-P values to processed rows"""
        for row in processed.rows:
            if row.article in sync_result.assigned_idp:
                row.id_p = sync_result.assigned_idp[row.article]

    async def _fill_idp_on_sheets(self, spreadsheet_id: str, config: Dict):
        """Fill ID-P on related sheets based on ID from primary (Главная)"""

        primary_sheet = config.get("sheets", {}).get("primary", "Главная")
        target_sheets = config.get("base_sheets_for_creation", [])

        if not target_sheets:
            return

        self._log(spreadsheet_id, "ПРАЙС", "Заполнение ID-P", f"Листы: {', '.join(target_sheets)}", "🔄")

        try:
            # 1. Read ID -> ID-P mapping from primary sheet
            ws_primary = self.sheets.get_worksheet(spreadsheet_id, primary_sheet)
            primary_values = ws_primary.get_all_values()

            if not primary_values:
                return

            primary_headers = primary_values[0]
            id_col = self._find_column_index(primary_headers, "ID", 0)
            idp_col = self._find_column_index(primary_headers, "ID-P", 1)

            if id_col == -1 or idp_col == -1:
                logger.warning("ID or ID-P column not found in primary sheet")
                return

            # Build ID -> ID-P map
            id_to_idp: Dict[str, str] = {}
            for row in primary_values[1:]:
                id_val = self._as_trimmed_string(self._safe_get(row, id_col))
                idp_val = self._as_trimmed_string(self._safe_get(row, idp_col))
                if id_val and idp_val:
                    id_to_idp[id_val] = idp_val

            # 2. Fill ID-P on each target sheet
            for sheet_name in target_sheets:
                try:
                    ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)
                    values = ws.get_all_values()

                    if not values:
                        continue

                    headers = values[0]
                    sheet_id_col = self._find_column_index(headers, "ID", -1)
                    sheet_idp_col = self._find_column_index(headers, "ID-P", -1)

                    if sheet_id_col == -1 or sheet_idp_col == -1:
                        logger.warning(f"ID or ID-P column not found in {sheet_name}")
                        continue

                    # Prepare batch updates
                    updates = []
                    idp_col_letter = self._col_index_to_letter(sheet_idp_col)

                    for row_idx, row in enumerate(values[1:], start=2):
                        id_val = self._as_trimmed_string(self._safe_get(row, sheet_id_col))
                        current_idp = self._as_trimmed_string(self._safe_get(row, sheet_idp_col))

                        if id_val and id_val in id_to_idp and not current_idp:
                            updates.append({
                                "range": f"'{sheet_name}'!{idp_col_letter}{row_idx}",
                                "values": [[id_to_idp[id_val]]]
                            })

                    if updates:
                        # Batch update in chunks to avoid quota limits
                        chunk_size = 500
                        for i in range(0, len(updates), chunk_size):
                            chunk = updates[i:i + chunk_size]
                            ws.batch_update(chunk, value_input_option="RAW")

                        logger.info(f"Filled {len(updates)} ID-P values on {sheet_name}")

                except Exception as e:
                    logger.warning(f"Failed to fill ID-P on {sheet_name}: {e}")

        except Exception as e:
            logger.error(f"Fill ID-P on sheets failed: {e}", exc_info=True)

    def _col_index_to_letter(self, col_idx: int) -> str:
        """Convert column index (0-based) to letter (A, B, ... Z, AA, AB, ...)"""
        result = ""
        while col_idx >= 0:
            result = chr(65 + (col_idx % 26)) + result
            col_idx = col_idx // 26 - 1
        return result

    async def _copy_prices_to_sheets(
        self,
        spreadsheet_id: str,
        processed: ProcessedData,
        config: Dict,
        price_multiplier: float = 1.0
    ):
        """Copy prices from processed data to price dynamics and calculation sheets"""

        # Target sheets for price copy
        target_sheets = [
            config.get("sheets", {}).get("price_dynamics", "Динамика цены"),
            config.get("sheets", {}).get("price_calculation", "Расчет цены")
        ]

        self._log(
            spreadsheet_id,
            "ПРАЙС",
            "Копирование цен",
            f"Листы: {', '.join(target_sheets)}, множитель: {price_multiplier}",
            "🔄"
        )

        # Build ID-P -> price map from processed data
        idp_to_price: Dict[str, float] = {}
        for row in processed.rows:
            if row.id_p and row.price > 0:
                final_price = row.price * price_multiplier if price_multiplier != 1.0 else row.price
                idp_to_price[row.id_p] = final_price

        if not idp_to_price:
            logger.info("No prices to copy")
            return

        for sheet_name in target_sheets:
            try:
                ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)
                values = ws.get_all_values()

                if not values:
                    continue

                headers = values[0]
                idp_col = self._find_column_index(headers, "ID-P", -1)
                price_col = self._find_column_index(headers, "ЦЕНА EXW", -1)

                # Also try alternative price column names
                if price_col == -1:
                    for alt_name in ["ЦЕНА EXW из Б/З", "EXW", "Цена"]:
                        price_col = self._find_column_index(headers, alt_name, -1)
                        if price_col >= 0:
                            break

                if idp_col == -1 or price_col == -1:
                    logger.warning(f"ID-P or Price column not found in {sheet_name}")
                    continue

                # Prepare batch updates
                updates = []
                price_col_letter = self._col_index_to_letter(price_col)

                for row_idx, row in enumerate(values[1:], start=2):
                    idp_val = self._as_trimmed_string(self._safe_get(row, idp_col))

                    if idp_val and idp_val in idp_to_price:
                        updates.append({
                            "range": f"'{sheet_name}'!{price_col_letter}{row_idx}",
                            "values": [[idp_to_price[idp_val]]]
                        })

                if updates:
                    # Batch update in chunks
                    chunk_size = 500
                    for i in range(0, len(updates), chunk_size):
                        chunk = updates[i:i + chunk_size]
                        ws.batch_update(chunk, value_input_option="USER_ENTERED")

                    logger.info(f"Copied {len(updates)} prices to {sheet_name}")

            except Exception as e:
                logger.warning(f"Failed to copy prices to {sheet_name}: {e}")

    async def _recalculate_formulas(self, spreadsheet_id: str, config: Dict):
        """
        Recalculate formulas on price dynamics and calculation sheets.

        This method triggers a refresh of calculated columns by reading and updating
        formula cells. Google Sheets automatically recalculates formulas when values
        they depend on change, so this mainly ensures the sheet is refreshed.
        """

        sheets_to_refresh = [
            config.get("sheets", {}).get("price_dynamics", "Динамика цены"),
            config.get("sheets", {}).get("price_calculation", "Расчет цены")
        ]

        self._log(spreadsheet_id, "ПРАЙС", "Пересчет формул", f"Листы: {', '.join(sheets_to_refresh)}", "🔄")

        # In Google Sheets, formulas recalculate automatically when dependent cells change
        # We just need to ensure the sheets are accessed to trigger any pending recalculations
        for sheet_name in sheets_to_refresh:
            try:
                ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)
                # Reading the first row triggers a refresh
                _ = ws.row_values(1)
                logger.info(f"Refreshed {sheet_name}")
            except Exception as e:
                logger.warning(f"Failed to refresh {sheet_name}: {e}")

    async def _update_statuses(self, spreadsheet_id: str, config: Dict):
        """
        Update statuses after processing (for samples mode).

        Updates product statuses based on ID-P presence:
        - If ID-P is filled and status is empty → set to "New в работу"
        """

        primary_sheet = config.get("sheets", {}).get("primary", "Главная")
        status_new = config.get("constants", {}).get("statuses", {}).get("new", "New в работу")

        self._log(spreadsheet_id, "ПРАЙС", "Обновление статусов", f"Лист: {primary_sheet}", "🔄")

        try:
            ws = self.sheets.get_worksheet(spreadsheet_id, primary_sheet)
            values = ws.get_all_values()

            if not values:
                return

            headers = values[0]
            idp_col = self._find_column_index(headers, "ID-P", -1)
            status_col = self._find_column_index(headers, "Статус", -1)

            if idp_col == -1 or status_col == -1:
                logger.warning("ID-P or Status column not found in primary sheet")
                return

            updates = []
            status_col_letter = self._col_index_to_letter(status_col)

            for row_idx, row in enumerate(values[1:], start=2):
                idp_val = self._as_trimmed_string(self._safe_get(row, idp_col))
                status_val = self._as_trimmed_string(self._safe_get(row, status_col))

                # If ID-P exists but status is empty, set to "New в работу"
                if idp_val and not status_val:
                    updates.append({
                        "range": f"'{primary_sheet}'!{status_col_letter}{row_idx}",
                        "values": [[status_new]]
                    })

            if updates:
                chunk_size = 500
                for i in range(0, len(updates), chunk_size):
                    chunk = updates[i:i + chunk_size]
                    ws.batch_update(chunk, value_input_option="RAW")

                logger.info(f"Updated {len(updates)} statuses")
                self._log(spreadsheet_id, "ПРАЙС", "Статусы обновлены", f"Записей: {len(updates)}", "✅")

        except Exception as e:
            logger.error(f"Update statuses failed: {e}", exc_info=True)


# Singleton instance
_price_processor: Optional[PriceProcessor] = None


def get_price_processor(sheets_service: SheetsService, logging_service=None) -> PriceProcessor:
    """Get or create PriceProcessor instance"""
    global _price_processor
    if _price_processor is None:
        _price_processor = PriceProcessor(sheets_service, logging_service)
    return _price_processor
