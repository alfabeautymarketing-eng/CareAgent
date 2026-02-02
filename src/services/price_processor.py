"""
Price Processor Service

Handles parsing and processing of supplier price lists (Б/З поставщик).
Replaces GAS functions: processMtMainPrice, processMtTesterPrice, processMtSamplesPrice,
processSkPriceSheet, processSkPriceProbes, processSsPriceSheet.

Full processing flow (matches original GAS logic):
1. Create snapshot in source document
2. Parse data from source
3. Clear ID-P, ID-G, and price columns
4. Sync with main sheet (Главная)
5. Fill ID-G for rows without ID-P
6. Fill ID-P on all sheets
7. Copy prices to related sheets
8. Apply formulas on Динамика цены and Расчет цены
9. Replicate new articles
10. Update statuses (samples mode only)
"""

import asyncio
import yaml
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field

from src.utils.logger import logger
from src.services.sheets import SheetsService
from src.services.product_matcher import ProductMatcher
from src.models.product import (
    Product, ProjectCode, SupplierInfo, LocalizationInfo, 
    PriceInfo, ProductStatus, ProductType
)
from src.storage.supabase_store import get_supabase_product_store


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
    new_articles: List[Dict[str, Any]] = field(default_factory=list)


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

    def __init__(self, sheets_service: SheetsService, sync_service: Any, logging_service=None):
        self.sheets = sheets_service
        self.sync_service = sync_service
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

    async def process_all(
        self,
        project: str,
        spreadsheet_id: str,
        source_doc_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Run ALL processing cycles for a project sequentially.

        Matches GAS behavior where each project has multiple processing stages:
        - MT: main (Б/З поставщик) → tester (Тестер) → samples (Пробники)
        - SK: main (Б/З поставщик) → probes (Пробники)
        - SS: main (Б/З поставщик)

        Each cycle runs to completion before the next one starts.
        """
        project = project.lower()
        config = self.load_config(project)

        cycles = config.get("processing_cycles", [])
        if not cycles:
            # Fallback: single main cycle
            cycles = [{"mode": "main"}]

        self._log(
            spreadsheet_id,
            "ПРАЙС",
            f"Полная обработка {project.upper()}",
            f"Циклов: {len(cycles)} ({', '.join(c['mode'] for c in cycles)})",
            "🚀"
        )

        results = []
        total_rows = 0
        total_new = 0

        for i, cycle in enumerate(cycles):
            cycle_mode = cycle["mode"]
            cycle_num = i + 1

            self._log(
                spreadsheet_id,
                "ПРАЙС",
                f"Цикл {cycle_num}/{len(cycles)}: {cycle_mode}",
                f"Старт",
                "🔄"
            )

            try:
                result = await self.process(
                    project=project,
                    mode=cycle_mode,
                    spreadsheet_id=spreadsheet_id,
                    source_doc_id=source_doc_id,
                    dry_run=False
                )
                results.append(result)
                total_rows += result.get("processed_rows", 0)
                total_new += result.get("new_articles", 0)

                self._log(
                    spreadsheet_id,
                    "ПРАЙС",
                    f"Цикл {cycle_num}/{len(cycles)}: {cycle_mode} завершён",
                    f"Строк: {result.get('processed_rows', 0)}",
                    "✅"
                )

                # Pause between cycles to avoid rate limiting (60 req/min quota)
                if i < len(cycles) - 1:
                    logger.info(f"Pausing 15s before next cycle to avoid rate limits...")
                    await asyncio.sleep(15)

            except Exception as e:
                logger.error(f"Cycle {cycle_mode} failed: {e}", exc_info=True)
                self._log(
                    spreadsheet_id,
                    "ПРАЙС",
                    f"Цикл {cycle_num}/{len(cycles)}: {cycle_mode} ОШИБКА",
                    str(e),
                    "❌"
                )
                results.append({
                    "status": "error",
                    "mode": cycle_mode,
                    "message": str(e),
                    "processed_rows": 0
                })

        self._log(
            spreadsheet_id,
            "ПРАЙС",
            f"Полная обработка {project.upper()} завершена",
            f"Циклов: {len(cycles)}, строк: {total_rows}, новых: {total_new}",
            "🏁"
        )

        return {
            "status": "success",
            "message": f"Полная обработка {project.upper()}: {len(cycles)} циклов, {total_rows} строк",
            "cycles": len(cycles),
            "total_rows": total_rows,
            "total_new": total_new,
            "results": results
        }

    async def process(
        self,
        project: str,
        mode: str,
        spreadsheet_id: str,
        source_doc_id: Optional[str] = None,
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Main processing method - FULL GAS LOGIC IMPLEMENTATION.

        Workflow (matches original GAS processMtMainPrice):
        1. Load config and parser settings
        2. Create snapshot in source document (NEW)
        3. Read source data
        4. Parse data based on project/mode
        5. If dry_run, return preview
        6. Clear columns: ID-P (5 sheets), ЦЕНА EXW, ID-G (main mode only)
        7. Sync with main sheet (Главная)
        8. Fill ID-G for rows without ID-P (NEW)
        9. Fill ID-P on ALL sheets (optimized)
        10. Copy prices to related sheets
        11. Apply REAL formulas on Динамика цены (NEW)
        12. Apply INDEX/MATCH formulas on Расчет цены (NEW)
        13. Replicate new articles
        14. Update statuses (samples mode only)

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

        self._log(spreadsheet_id, "ПРАЙС", f"Обработка {project.upper()} {mode}", "Старт (полная логика GAS)", "🚀")

        try:
            # 1. Load config
            config = self.load_config(project)
            parser_config = self.get_parser_config(project, mode)

            # Get source document ID
            actual_source_doc_id = source_doc_id
            if not actual_source_doc_id:
                source_config = config.get("source", {})
                actual_source_doc_id = source_config.get("doc_id", spreadsheet_id)

            # Get source sheet name - priority order:
            # 1. processing_cycles[].source_sheet (from YAML cycles config)
            # 2. source.sheets.{mode} (from YAML source section)
            # 3. fallback to "-Б/З поставщик"
            # NOTE: parser_config.sheet_name is parser description, NOT the actual source sheet
            cycles = config.get("processing_cycles", [])
            cycle_config = next((c for c in cycles if c.get("mode") == mode), {})
            source_sheet_name = cycle_config.get("source_sheet")
            if not source_sheet_name:
                source_sheets = config.get("source", {}).get("sheets", {})
                source_sheet_name = source_sheets.get(mode, "-Б/З поставщик")

            # 2. Create snapshot in source document (NEW - matches GAS)
            if not dry_run and actual_source_doc_id:
                try:
                    await self._create_source_snapshot(
                        actual_source_doc_id,
                        source_sheet_name,
                        spreadsheet_id
                    )
                except Exception as e:
                    logger.warning(f"Snapshot creation failed (non-critical): {e}")

            # 3. Read source data
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

            # 4. Parse data based on project/mode
            processed = self._parse_data(
                source_data["values"],
                project,
                mode,
                parser_config,
                source_data.get("backgrounds")
            )

            if not processed.rows:
                return {
                    "status": "error",
                    "message": "Не найдено данных для обработки",
                    "processed_rows": 0
                }

            # 5. If dry_run, return preview
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

            # 6. CLEAR operations (only for main mode - tester/samples continue numbering)
            if mode == "main":
                self._log(spreadsheet_id, "ПРАЙС", "Очистка столбцов", "ID-P, ЦЕНА EXW, ID-G", "🧹")
                # Clear ID-P on all 5 sheets
                await self._clear_idp_columns(spreadsheet_id, config)
                # Clear ЦЕНА EXW из Б/З
                await self._clear_price_exw_columns(spreadsheet_id, config)
                # Clear ID-G on main sheet (NEW)
                await self._clear_idg_column_on_main(spreadsheet_id, config)

            # 7. Sync with main sheet (Главная)
            sync_result = await self._sync_with_main(
                spreadsheet_id,
                processed,
                config,
                project,
                product_matcher=ProductMatcher(spreadsheet_id)
            )

            # 8. Fill ID-G for rows without ID-P (NEW - matches GAS _fillIdgForRowsWithoutIdp)
            await self._fill_idg_for_rows_without_idp(spreadsheet_id, config)

            # 9. Apply assigned ID-P to processed data
            self._apply_assigned_idp(processed, sync_result)

            # Build ID → ID-P mapping for optimized fill
            id_to_idp: Dict[str, str] = {}
            try:
                primary_sheet = config.get("sheets", {}).get("primary", "Главная")
                ws = self.sheets.get_worksheet(spreadsheet_id, primary_sheet)
                values = ws.get_all_values()
                if values:
                    headers = values[0]
                    id_col = self._find_column_index(headers, "ID", -1)
                    idp_col = self._find_column_index(headers, "ID-P", -1)
                    if id_col >= 0 and idp_col >= 0:
                        for row in values[1:]:
                            id_val = self._as_trimmed_string(self._safe_get(row, id_col))
                            idp_val = self._as_trimmed_string(self._safe_get(row, idp_col))
                            if id_val and idp_val:
                                id_to_idp[id_val] = idp_val
            except Exception as e:
                logger.warning(f"Failed to build ID→ID-P mapping: {e}")

            # 10. Fill ID-P on ALL sheets (OPTIMIZED - single batch)
            await self._fill_idp_on_all_sheets_optimized(spreadsheet_id, config, id_to_idp)

            # 11. Copy prices to related sheets
            price_multiplier = parser_config.get("price_multiplier", 1)
            await self._copy_prices_to_sheets(spreadsheet_id, processed, config, price_multiplier)

            # 12. Apply REAL formulas on Динамика цены (NEW - not just refresh)
            await self._apply_price_dynamics_formulas(spreadsheet_id, config)

            # 13. Apply INDEX/MATCH formulas on Расчет цены (NEW)
            await self._apply_price_calculation_formulas(spreadsheet_id, config)

            # 14. Replicate new articles across all sheets
            if sync_result.new_articles:
                self._log(
                    spreadsheet_id,
                    "АРТИКУЛ",
                    "Синхронизация новых артикулов",
                    f"Количество: {len(sync_result.new_articles)}",
                    "🔄"
                )
                await self.sync_service.replicate_new_articles(spreadsheet_id, sync_result.new_articles)

            # 15. Update statuses (only for samples mode)
            if mode == "samples" and parser_config.get("update_statuses_after"):
                await self._update_statuses(spreadsheet_id, config)

            self._log(
                spreadsheet_id,
                "ПРАЙС",
                f"Обработка {project.upper()} {mode} завершена",
                f"Строк: {len(processed.rows)}, групп: {processed.unique_groups}, новых: {len(sync_result.created_rows)}",
                "✅"
            )

            return {
                "status": "success",
                "message": f"Обработано {len(processed.rows)} артикулов (полная логика GAS)",
                "processed_rows": len(processed.rows),
                "groups_found": processed.unique_groups,
                "new_articles": len(sync_result.created_rows),
                "updated_articles": len(sync_result.updated_rows),
                "barcode_mismatches": len(sync_result.barcode_mismatches),
                "idp_filled": len(id_to_idp),
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
            logger.info(f"Reading source data: doc={source_doc_id}, sheet='{sheet_name}', mode={mode}")
            ws = self.sheets.get_worksheet(source_doc_id, sheet_name)
            values = ws.get_all_values()
            logger.info(f"Source sheet '{sheet_name}': {len(values)} rows, {len(values[0]) if values else 0} cols")
            if values and len(values) > 1:
                logger.info(f"Source sheet '{sheet_name}' row[0] (headers): {values[0][:8]}")
                logger.info(f"Source sheet '{sheet_name}' row[1] (first data): {values[1][:8]}")

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
            logger.error(f"Failed to read source data from sheet '{sheet_name}': {e}")
            raise ValueError(f"Не удалось прочитать данные из листа '{sheet_name}': {e}")

    def _parse_data(
        self,
        values: List[List[Any]],
        project: str,
        mode: str,
        parser_config: Dict,
        backgrounds: Optional[List[List[str]]] = None
    ) -> ProcessedData:
        """Parse source data based on project and mode"""

        if project == "mt":
            return self._parse_mt_data(values, mode, parser_config)
        elif project == "sk":
            return self._parse_sk_data(values, mode, parser_config, backgrounds)
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
        logger.info(f"_parse_mt_data mode={mode}: header_row_index={header_row_index}, total_rows={len(values)}")
        if header_row_index == -1:
            raise ValueError("Не найдена строка заголовков с CODE и DESCRIPTION")

        headers = values[header_row_index]
        logger.info(f"_parse_mt_data mode={mode}: headers={headers[:8]}")

        # Get column indices
        excel_cols = parser_config.get("excel_columns", self.DEFAULT_EXCEL_COLUMNS)
        code_idx = self._find_column_index(headers, "CODE", excel_cols.get("CODE", 0))
        # CATEGORY is optional (e.g. -Тестер sheet has no CATEGORY column)
        # Pass default=-1 so it returns -1 when not found, instead of falling back to index 1
        category_idx = self._find_column_index(headers, "CATEGORY", -1)
        desc_idx = self._find_column_index(headers, "DESCRIPTION", excel_cols.get("DESCRIPTION", 2))
        format_idx = self._find_column_index(headers, "FORMAT", excel_cols.get("FORMAT", 3))
        units_idx = self._find_column_index(headers, "UNITS", excel_cols.get("UNITS", 4))
        price_idx = self._find_column_index(headers, "PRICE", excel_cols.get("PRICE", 5))
        ean_idx = self._find_column_index(headers, "EAN", excel_cols.get("EAN", 6))

        logger.info(f"_parse_mt_data mode={mode}: code_idx={code_idx}, category_idx={category_idx}, desc_idx={desc_idx}, price_idx={price_idx}")

        if code_idx == -1 or desc_idx == -1:
            raise ValueError("Не найдены обязательные столбцы CODE и DESCRIPTION")

        # Get transformations for mode
        transformations = parser_config.get("transformations", {})
        volume_prefix = transformations.get("volume_prefix", "")
        group_suffix = transformations.get("group_suffix", "")
        price_multiplier = parser_config.get("price_multiplier", 1)

        # Parse rows
        skipped_no_code = 0
        skipped_no_desc = 0
        group_rows = 0
        for i in range(header_row_index + 1, len(values)):
            row = values[i]
            code_value = self._as_trimmed_string(self._safe_get(row, code_idx))
            category_value = self._as_trimmed_string(self._safe_get(row, category_idx)) if category_idx != -1 else ""
            desc_value = self._as_trimmed_string(self._safe_get(row, desc_idx))

            if not code_value:
                skipped_no_code += 1
                continue

            # Group detection: CODE exists AND CATEGORY empty
            if code_value and not category_value and not desc_value:
                current_group = code_value
                group_rows += 1
                continue

            if not desc_value:
                skipped_no_desc += 1
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

        logger.info(
            f"_parse_mt_data mode={mode} result: "
            f"parsed={len(rows)}, groups={group_rows}, "
            f"skipped_no_code={skipped_no_code}, skipped_no_desc={skipped_no_desc}, "
            f"total_data_rows={len(values) - header_row_index - 1}"
        )

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
        parser_config: Dict,
        backgrounds: Optional[List[List[str]]] = None
    ) -> ProcessedData:
        """
        Parse SK project data with color-based group detection.

        SK specifics:
        - Groups detected by yellow background color (#ffff00)
        - Combined group format: "{line} - {group}"
        - RRP column present
        - Probes mode has different logic (not_starts_with_00)
        """

        rows: List[ParsedRow] = []
        articles: List[str] = []
        groups: List[str] = []
        current_group = ""
        current_line = ""

        # Group detection config
        group_detection = parser_config.get("group_detection", {})
        group_color = group_detection.get("group_color", "#ffff00").lower()

        # Find header row
        header_row_index = self._find_header_row_sk(values)
        if header_row_index == -1:
            # Fallback to MT-style parsing
            logger.warning("SK header row not found, falling back to MT parsing")
            return self._parse_mt_data(values, mode, parser_config)

        headers = values[header_row_index]

        # Get column indices from headers
        code_col = self._find_column_index(headers, "CODE", -1)
        product_col = self._find_column_index(headers, "PRODUCT", -1)
        units_col = self._find_column_index(headers, "UNITS", -1)
        price_col = self._find_column_index(headers, "PRICE", -1)
        rrp_col = self._find_column_index(headers, "RRP", -1)

        if code_col == -1 or product_col == -1:
            logger.warning("Required columns not found in SK sheet")
            return self._parse_mt_data(values, mode, parser_config)

        # Parse rows
        for i in range(header_row_index + 1, len(values)):
            row = values[i]

            # Check background color for group detection
            is_group_row = False
            if backgrounds and i < len(backgrounds):
                row_backgrounds = backgrounds[i]
                # Check if first cell has group color
                if row_backgrounds and len(row_backgrounds) > 0:
                    cell_color = row_backgrounds[0].lower() if row_backgrounds[0] else ""
                    is_group_row = self._colors_match(cell_color, group_color)

            code_value = self._as_trimmed_string(self._safe_get(row, code_col))
            product_value = self._as_trimmed_string(self._safe_get(row, product_col))

            # Group detection by color
            if is_group_row and product_value:
                # Parse group: could be "Line - Group" or just "Group"
                if " - " in product_value:
                    parts = product_value.split(" - ", 1)
                    current_line = parts[0].strip()
                    current_group = parts[1].strip() if len(parts) > 1 else ""
                else:
                    current_group = product_value
                continue

            # Article detection: has code and product name
            if code_value and product_value:
                units_value = self._get_value(row, units_col)
                price_value = self._parse_price(self._get_value(row, price_col))
                rrp_value = self._get_value(row, rrp_col) if rrp_col >= 0 else ""

                # Build combined group
                combined_group = f"{current_line} - {current_group}" if current_line else current_group

                parsed_row = ParsedRow(
                    article=code_value,
                    name_eng=product_value,
                    volume="",  # SK doesn't have separate volume column
                    barcode="",  # SK doesn't have barcode in price list
                    units_per_pack=str(units_value) if units_value else "",
                    price=price_value,
                    group=combined_group
                )

                rows.append(parsed_row)
                articles.append(code_value)
                groups.append(combined_group)

        return ProcessedData(
            headers=self.OUTPUT_HEADERS.copy(),
            rows=rows,
            articles=articles,
            groups=groups
        )

    def _find_header_row_sk(self, values: List[List[Any]], max_rows: int = 10) -> int:
        """Find row index containing CODE and PRODUCT headers for SK"""
        for i in range(min(len(values), max_rows)):
            row = values[i]
            has_code = any("CODE" in str(cell).upper() for cell in row)
            has_product = any("PRODUCT" in str(cell).upper() for cell in row)
            if has_code and has_product:
                return i
        return -1

    def _colors_match(self, color1: str, color2: str) -> bool:
        """Check if two colors match (handling different formats)"""
        # Normalize colors
        c1 = color1.lower().replace("#", "").strip()
        c2 = color2.lower().replace("#", "").strip()

        if not c1 or not c2:
            return False

        # Direct match
        if c1 == c2:
            return True

        # Handle RGB format vs hex
        # ffff00 = yellow
        if c1 in ["ffff00", "yellow"] and c2 in ["ffff00", "yellow"]:
            return True

        return False

    def _parse_ss_data(
        self,
        values: List[List[Any]],
        mode: str,
        parser_config: Dict
    ) -> ProcessedData:
        """
        Parse SS project data with markers (-ПРОФ, SAMPLES).

        SS specifics:
        - "PROFESSIONAL PRODUCTS" marker adds "-ПРОФ" suffix to group
        - "PROMOTIONAL MATERIALS" marker stops processing
        - SAMPLES group: extract number from "form" to "units_per_pack"
        """

        rows: List[ParsedRow] = []
        articles: List[str] = []
        groups: List[str] = []
        current_group = ""
        is_professional_mode = False

        # Get column config
        excel_cols = parser_config.get("excel_columns", {})
        code_idx = excel_cols.get("CODE", 1)
        name_idx = excel_cols.get("PRODUCT_NAME", 2)
        size_idx = excel_cols.get("SIZE", 3)
        pack_idx = excel_cols.get("PACK", 4)
        barcode_idx = excel_cols.get("BAR_CODE_ACL", 5)
        qty_idx = excel_cols.get("QTY_BOX", 6)
        price_idx = excel_cols.get("EX_WORKS_CARROS", 7)

        # Get markers config
        markers = parser_config.get("markers", {})
        prof_trigger = markers.get("professional_mode", {}).get("trigger", "PROFESSIONAL PRODUCTS")
        prof_suffix = markers.get("professional_mode", {}).get("group_suffix", "-ПРОФ")
        stop_trigger = markers.get("stop_processing", {}).get("trigger", "PROMOTIONAL MATERIALS")

        # SAMPLES logic config
        samples_config = parser_config.get("samples_logic", {})
        samples_enabled = samples_config.get("enabled", False)
        samples_group_name = samples_config.get("group_name", "SAMPLES")

        # Find header row
        header_row_index = self._find_header_row_ss(values)

        # Parse rows
        for i in range(header_row_index + 1, len(values)):
            row = values[i]

            # Get first cell value for marker detection
            first_cell = self._as_trimmed_string(self._safe_get(row, 0)).upper()

            # Check for stop marker
            if stop_trigger.upper() in first_cell:
                logger.info(f"SS: Stop marker found at row {i + 1}")
                break

            # Check for professional mode marker
            if prof_trigger.upper() in first_cell:
                is_professional_mode = True
                logger.info(f"SS: Professional mode enabled at row {i + 1}")
                continue

            # Get values
            code_value = self._as_trimmed_string(self._safe_get(row, code_idx))
            name_value = self._as_trimmed_string(self._safe_get(row, name_idx))
            size_value = self._get_value(row, size_idx)
            pack_value = self._get_value(row, pack_idx)
            barcode_value = self._get_value(row, barcode_idx)
            qty_value = self._get_value(row, qty_idx)
            price_value = self._parse_price(self._get_value(row, price_idx))

            # Group detection: has name but no code (or specific pattern)
            if name_value and not code_value:
                current_group = name_value
                # Add professional suffix if in professional mode
                if is_professional_mode:
                    current_group = f"{current_group}{prof_suffix}"
                continue

            # Article detection: has code and name
            if code_value and name_value:
                # Determine final group
                group = current_group

                # SAMPLES special logic: extract number from pack/form
                units = qty_value
                if samples_enabled and current_group.upper() == samples_group_name:
                    # Try to extract number from pack_value
                    extracted = re.search(r'\d+', pack_value)
                    if extracted:
                        units = extracted.group()

                parsed_row = ParsedRow(
                    article=code_value,
                    name_eng=name_value,
                    volume=size_value,
                    barcode=str(barcode_value) if barcode_value else "",
                    units_per_pack=str(units) if units else "",
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

    def _find_header_row_ss(self, values: List[List[Any]], max_rows: int = 10) -> int:
        """Find row index containing CODE or PRODUCT headers for SS"""
        for i in range(min(len(values), max_rows)):
            row = values[i]
            # Look for typical SS headers
            row_text = " ".join(str(cell).upper() for cell in row)
            if "CODE" in row_text or "PRODUCT" in row_text or "SIZE" in row_text:
                return i
        return 0  # Default to first row

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

        for idx, sheet_name in enumerate(sheets_to_clear):
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

                # Delay between sheets to avoid rate limits
                if idx < len(sheets_to_clear) - 1:
                    await asyncio.sleep(2)

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
        project: str,
        product_matcher: Optional[ProductMatcher] = None
    ) -> SyncResult:
        """Sync processed data with main sheet (Главная)"""

        result = SyncResult()
        primary_sheet = config.get("sheets", {}).get("primary", "Главная")
        # ID-P = plain numbers (1, 2, 3...) — NO prefix (matches GAS behavior)
        # ID = project prefix for display (MT-001, SK-001...)
        id_prefix = config.get("project", {}).get("code", project).upper() + "-"

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
            
            # Additional columns for Supabase sync
            idg_col = self._find_column_index(headers, "ID-G", -1)
            idl_col = self._find_column_index(headers, "ID-L", -1)
            
            nameru_col = self._find_column_index(headers, "Название (рус)", -1)
            if nameru_col == -1:
                nameru_col = self._find_column_index(headers, "Название RUS", -1)
                
            line_col = self._find_column_index(headers, "Линия", -1)
            if line_col == -1:
                line_col = self._find_column_index(headers, "Линейка", -1)

            # Prepare for Supabase sync
            products_to_sync: List[Product] = []
            try:
                project_enum = ProjectCode(project.lower())
            except ValueError:
                project_enum = ProjectCode.MT # Default fallback


            # Build article -> row index map
            article_map: Dict[str, int] = {}
            for i, row in enumerate(all_values[1:], start=2):
                art = self._as_trimmed_string(self._safe_get(row, article_col))
                if art:
                    article_map[art] = i

            # Find max numbers for IDs
            max_idp = 0
            max_id = 0
            for row in all_values[1:]:
                # ID-P — plain numbers (1, 2, 3...)
                idp_val = self._as_trimmed_string(self._safe_get(row, idp_col))
                if idp_val:
                    try:
                        num = int(idp_val)
                        max_idp = max(max_idp, num)
                    except ValueError:
                        # Try stripping old prefix if any legacy data
                        if "-" in idp_val:
                            try:
                                num = int(idp_val.split("-")[-1])
                                max_idp = max(max_idp, num)
                            except ValueError:
                                pass

                # ID — has project prefix (MT-001, SK-001...)
                id_val = self._as_trimmed_string(self._safe_get(row, id_col))
                if id_val and id_val.startswith(id_prefix):
                    try:
                        num = int(id_val[len(id_prefix):])
                        max_id = max(max_id, num)
                    except ValueError:
                        pass

            # Process each parsed row
            updates = []
            new_rows = []

            # Fetch base products ONCE before loop (avoid N+1 API calls)
            base_candidates = None
            if product_matcher:
                base_candidates = product_matcher.fetch_base_products()
                if base_candidates:
                    self._log(spreadsheet_id, "Smart Match", f"Loaded {len(base_candidates)} base products for matching", "", "📦")

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

                    # Assign ID-P if empty (plain number, no prefix)
                    existing_idp = self._as_trimmed_string(self._safe_get(existing_row, idp_col))
                    if not existing_idp:
                        max_idp += 1
                        new_idp = str(max_idp)
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
                    new_idp = str(max_idp)  # ID-P = plain number

                    max_id += 1
                    new_id = f"{id_prefix}{max_id:03d}"  # ID = MT-001 (with prefix)
                    
                    result.assigned_idp[article] = new_idp

                    # 🎯 SMART MATCH for new article
                    match_info = None
                    if product_matcher and parsed_row.name_eng and base_candidates:
                        self._log(spreadsheet_id, "Smart Match", f"Поиск для нового товара: {parsed_row.name_eng}", "", "🔍")
                        match_res = product_matcher.find_best_match(parsed_row.name_eng, candidates=base_candidates)
                        if match_res and match_res.get("match_found") and match_res.get("confidence", 0) >= 80:
                            match_info = match_res
                            self._log(spreadsheet_id, "Smart Match", f"Найдено соответствие ({match_res['confidence']}%): {match_res.get('best_match_id')}", f"Цель: {parsed_row.name_eng}", "✅")
                        else:
                            self._log(spreadsheet_id, "Smart Match", "Соответствие не найдено или низкая уверенность", f"Цель: {parsed_row.name_eng}", "ℹ️")

                    # Detect new_row initialization
                    new_row = [""] * len(headers)
                    if id_col >= 0:
                        new_row[id_col] = new_id
                    if idp_col >= 0:
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
                    
                    # Add to new_articles for further sheet replication
                    result.new_articles.append({
                        "id": new_id,
                        "idp": new_idp,
                        "article": article,
                        "name_eng": parsed_row.name_eng,
                        "match_info": match_info
                    })

                # --- Create Product object for Supabase sync ---
                try:
                    # Determine values
                    final_idp = new_idp if 'new_idp' in locals() else existing_idp

                    if final_idp:
                        # Values from existing row (if update) or empty (if new)
                        val_idg = ""
                        val_idl = ""
                        val_nameru = ""
                        val_line = ""

                        if 'existing_row' in locals():
                            val_idg = self._as_trimmed_string(self._safe_get(existing_row, idg_col)) if idg_col >= 0 else ""
                            val_idl = self._as_trimmed_string(self._safe_get(existing_row, idl_col)) if idl_col >= 0 else ""
                            val_nameru = self._as_trimmed_string(self._safe_get(existing_row, nameru_col)) if nameru_col >= 0 else ""
                            val_line = self._as_trimmed_string(self._safe_get(existing_row, line_col)) if line_col >= 0 else ""

                        # If smart match found something (for new rows), override NameRU
                        if 'match_info' in locals() and match_info and match_info.get("match_found"):
                             details = match_info.get("matched_product_details", {})
                             if details.get("Наименования рус по ДС"):
                                 val_nameru = details.get("Наименования рус по ДС")

                        # Supplier info from parsed row (always available)
                        sup_units = 1
                        if parsed_row.units_per_pack and str(parsed_row.units_per_pack).isdigit():
                             sup_units = int(parsed_row.units_per_pack)
                             
                        supplier_info = SupplierInfo(
                            article=parsed_row.article,
                            name_original=parsed_row.name_eng,
                            barcode=parsed_row.barcode,
                            units_per_pack=sup_units,
                            group=parsed_row.group,
                            line=val_line
                        )

                        prod = Product(
                            id=final_idp,
                            id_g=val_idg,
                            id_l=val_idl,
                            project=project_enum,
                            supplier=supplier_info,
                            localization=LocalizationInfo(name_ru=val_nameru, name_en=parsed_row.name_eng),
                            price=PriceInfo(base_price=float(parsed_row.price)),
                            volume=parsed_row.volume,
                            status=ProductStatus.ACTIVE,
                            product_type=ProductType.MAIN
                        )
                        products_to_sync.append(prod)

                except Exception as ex_prod:
                    logger.warning(f"Failed to create Product object for {parsed_row.article}: {ex_prod}")
                
                # Cleanup locals for next iteration
                if 'new_idp' in locals(): del new_idp
                if 'existing_row' in locals(): del existing_row
                if 'new_id' in locals(): del new_id


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
                                "range": f"{idp_col_letter}{row_idx}",
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
                            "range": f"{price_col_letter}{row_idx}",
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
                        "range": f"{status_col_letter}{row_idx}",
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

    # ============================================================================
    # NEW METHODS - Full GAS Logic Implementation
    # ============================================================================

    async def _create_source_snapshot(
        self,
        source_doc_id: str,
        sheet_name: str,
        spreadsheet_id: str
    ) -> str:
        """
        Create a snapshot of the source sheet with current date.
        Matches GAS behavior: creates "{sheet_name} {dd.MM.yy}" in source document.

        Args:
            source_doc_id: Source document ID
            sheet_name: Base sheet name (e.g., "-Б/З поставщик")
            spreadsheet_id: Target spreadsheet ID (for logging)

        Returns:
            Name of created snapshot sheet
        """
        try:
            # Format date like GAS: dd.MM.yy
            date_str = datetime.now().strftime("%d.%m.%y")
            snapshot_name = f"{sheet_name} {date_str}"

            self._log(spreadsheet_id, "ПРАЙС", "Создание snapshot", f"Лист: {snapshot_name}", "📸")

            # Get source spreadsheet
            source_ss = self.sheets.open_by_key(source_doc_id)

            # Read original sheet data
            original_ws = source_ss.worksheet(sheet_name)
            original_values = original_ws.get_all_values()

            # Delete old snapshots with same base name
            all_sheets = source_ss.worksheets()
            for ws in all_sheets:
                ws_name = ws.title
                if ws_name.startswith(f"{sheet_name} ") and ws_name != snapshot_name:
                    try:
                        source_ss.del_worksheet(ws)
                        logger.info(f"Deleted old snapshot: {ws_name}")
                    except Exception:
                        pass

            # Create or get snapshot sheet
            try:
                snapshot_ws = source_ss.worksheet(snapshot_name)
                snapshot_ws.clear()
            except Exception:
                snapshot_ws = source_ss.add_worksheet(
                    title=snapshot_name,
                    rows=max(len(original_values), 100),
                    cols=max(len(original_values[0]) if original_values else 10, 26)
                )

            # Copy data to snapshot
            if original_values:
                snapshot_ws.update(
                    range_name='A1',
                    values=original_values,
                    value_input_option='RAW'
                )

            logger.info(f"Created snapshot: {snapshot_name}")
            return snapshot_name

        except Exception as e:
            logger.warning(f"Failed to create snapshot: {e}")
            return sheet_name  # Return original name if snapshot fails

    async def _clear_idg_column_on_main(self, spreadsheet_id: str, config: Dict):
        """
        Clear ID-G column on main sheet (Главная).
        Matches GAS: _clearIdgColumnOnMain()
        """
        primary_sheet = config.get("sheets", {}).get("primary", "Главная")

        self._log(spreadsheet_id, "ПРАЙС", "Очистка ID-G", f"Лист: {primary_sheet}", "🔄")

        try:
            ws = self.sheets.get_worksheet(spreadsheet_id, primary_sheet)
            headers = ws.row_values(1)

            idg_col = self._find_column_index(headers, "ID-G", -1)
            if idg_col < 0:
                logger.warning("ID-G column not found in primary sheet")
                return

            num_rows = ws.row_count
            if num_rows > 1:
                col_letter = self._col_index_to_letter(idg_col)
                ws.batch_clear([f"{col_letter}2:{col_letter}{num_rows}"])
                logger.info(f"Cleared ID-G column on {primary_sheet}")

        except Exception as e:
            logger.warning(f"Failed to clear ID-G on main: {e}")

    async def _fill_idg_for_rows_without_idp(self, spreadsheet_id: str, config: Dict):
        """
        Fill ID-G for rows where ID-P is empty.
        Matches GAS: _fillIdgForRowsWithoutIdp()

        Logic:
        - Build map of group → ID-G from existing data
        - For rows without ID-P: assign ID-G based on group
        - Create new ID-G values for new groups
        """
        primary_sheet = config.get("sheets", {}).get("primary", "Главная")

        self._log(spreadsheet_id, "ПРАЙС", "Заполнение ID-G", f"Для строк без ID-P", "🔄")

        try:
            ws = self.sheets.get_worksheet(spreadsheet_id, primary_sheet)
            values = ws.get_all_values()

            if not values:
                return

            headers = values[0]
            idp_col = self._find_column_index(headers, "ID-P", -1)
            idg_col = self._find_column_index(headers, "ID-G", -1)
            group_col = self._find_column_index(headers, "Группа", -1)

            if idg_col < 0 or group_col < 0:
                logger.warning("ID-G or Группа column not found")
                return

            # Build map group → ID-G from existing data
            group_to_idg: Dict[str, str] = {}
            max_idg = 0

            for row in values[1:]:
                idg = self._as_trimmed_string(self._safe_get(row, idg_col))
                group = self._as_trimmed_string(self._safe_get(row, group_col))

                if idg and group:
                    group_to_idg[group] = idg
                    try:
                        num = int(idg)
                        max_idg = max(max_idg, num)
                    except ValueError:
                        pass

            # Fill ID-G for rows without ID-P
            updates = []
            idg_letter = self._col_index_to_letter(idg_col)

            for row_idx, row in enumerate(values[1:], start=2):
                idp = self._as_trimmed_string(self._safe_get(row, idp_col)) if idp_col >= 0 else ""
                idg = self._as_trimmed_string(self._safe_get(row, idg_col))
                group = self._as_trimmed_string(self._safe_get(row, group_col))

                # Fill ID-G only if: no ID-P, no ID-G, but has group
                if not idp and not idg and group:
                    if group in group_to_idg:
                        new_idg = group_to_idg[group]
                    else:
                        max_idg += 1
                        new_idg = str(max_idg)
                        group_to_idg[group] = new_idg

                    updates.append({
                        "range": f"{idg_letter}{row_idx}",
                        "values": [[new_idg]]
                    })

            if updates:
                # Batch update in chunks
                chunk_size = 500
                for i in range(0, len(updates), chunk_size):
                    chunk = updates[i:i + chunk_size]
                    ws.batch_update(chunk, value_input_option="RAW")

                logger.info(f"Filled {len(updates)} ID-G values")
                self._log(spreadsheet_id, "ПРАЙС", "ID-G заполнен", f"Записей: {len(updates)}", "✅")

        except Exception as e:
            logger.error(f"Fill ID-G failed: {e}", exc_info=True)

    def _find_year_column(self, headers: List[str], prefix: str, year: int) -> int:
        """Find column index for a year-specific header like 'EXW 2025, €'"""
        year_str = str(year)
        for i, h in enumerate(headers):
            h_upper = str(h or "").upper()
            if prefix.upper() in h_upper and year_str in h_upper:
                return i
        return -1

    async def _apply_price_dynamics_formulas(self, spreadsheet_id: str, config: Dict):
        """
        Apply real formulas on 'Динамика цены' sheet.
        Matches GAS: Lib.recalculatePriceDynamicsFormulas()

        Formulas:
        - EXW ALFASPA = EXW * (1 - СКИДКА)
        - Закупочная цена = EXW ALFASPA * курс
        - DDP-МОСКВА = Закупочная * коэффициент
        - Прирост = (текущий - предыдущий) / предыдущий
        """
        sheet_name = config.get("sheets", {}).get("price_dynamics", "Динамика цены")

        self._log(spreadsheet_id, "ПРАЙС", "Применение формул", f"Лист: {sheet_name}", "📊")

        try:
            ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)
            values = ws.get_all_values()

            if not values or len(values) < 2:
                return

            headers = values[0]
            current_year = datetime.now().year

            # Get exchange rate and DDP coefficient from config
            exchange_rate = config.get("constants", {}).get("exchange_rate", 100)
            ddp_coefficient = config.get("constants", {}).get("ddp_coefficient", 1.5)

            # Find year-specific columns
            exw_col = self._find_year_column(headers, "EXW", current_year)
            discount_col = self._find_year_column(headers, "СКИДКА ОТ EXW", current_year)
            exw_alfaspa_col = self._find_year_column(headers, "EXW ALFASPA", current_year)
            purchase_col = self._find_year_column(headers, "Закупочная цена", current_year)
            ddp_col = self._find_year_column(headers, "DDP-МОСКВА", current_year)

            # Also try without year suffix for growth columns
            growth_exw_col = self._find_column_index(headers, "Прирост EXW", -1)
            growth_ddp_col = self._find_column_index(headers, "Прирост DDP", -1)

            # Find previous year columns for growth calculation
            prev_year = current_year - 1
            prev_exw_alfaspa_col = self._find_year_column(headers, "EXW ALFASPA", prev_year)
            prev_ddp_col = self._find_year_column(headers, "DDP-МОСКВА", prev_year)

            updates = []
            num_rows = len(values)

            for row_idx in range(2, num_rows + 1):
                # EXW ALFASPA = EXW * (1 - СКИДКА)
                if exw_col >= 0 and discount_col >= 0 and exw_alfaspa_col >= 0:
                    exw_letter = self._col_index_to_letter(exw_col)
                    discount_letter = self._col_index_to_letter(discount_col)
                    alfaspa_letter = self._col_index_to_letter(exw_alfaspa_col)
                    formula = f'=IF({exw_letter}{row_idx}<>"",{exw_letter}{row_idx}*(1-{discount_letter}{row_idx}),"")'
                    updates.append({
                        "range": f"{alfaspa_letter}{row_idx}",
                        "values": [[formula]]
                    })

                # Закупочная цена = EXW ALFASPA * курс
                if exw_alfaspa_col >= 0 and purchase_col >= 0:
                    alfaspa_letter = self._col_index_to_letter(exw_alfaspa_col)
                    purchase_letter = self._col_index_to_letter(purchase_col)
                    formula = f'=IF({alfaspa_letter}{row_idx}<>"",{alfaspa_letter}{row_idx}*{exchange_rate},"")'
                    updates.append({
                        "range": f"{purchase_letter}{row_idx}",
                        "values": [[formula]]
                    })

                # DDP-МОСКВА = Закупочная * коэффициент
                if purchase_col >= 0 and ddp_col >= 0:
                    purchase_letter = self._col_index_to_letter(purchase_col)
                    ddp_letter = self._col_index_to_letter(ddp_col)
                    formula = f'=IF({purchase_letter}{row_idx}<>"",{purchase_letter}{row_idx}*{ddp_coefficient},"")'
                    updates.append({
                        "range": f"{ddp_letter}{row_idx}",
                        "values": [[formula]]
                    })

                # Прирост EXW = (текущий - предыдущий) / предыдущий
                if growth_exw_col >= 0 and exw_alfaspa_col >= 0 and prev_exw_alfaspa_col >= 0:
                    curr_letter = self._col_index_to_letter(exw_alfaspa_col)
                    prev_letter = self._col_index_to_letter(prev_exw_alfaspa_col)
                    growth_letter = self._col_index_to_letter(growth_exw_col)
                    formula = f'=IF(AND({curr_letter}{row_idx}<>"",{prev_letter}{row_idx}<>"",{prev_letter}{row_idx}<>0),({curr_letter}{row_idx}-{prev_letter}{row_idx})/{prev_letter}{row_idx},"")'
                    updates.append({
                        "range": f"{growth_letter}{row_idx}",
                        "values": [[formula]]
                    })

                # Прирост DDP = (текущий - предыдущий) / предыдущий
                if growth_ddp_col >= 0 and ddp_col >= 0 and prev_ddp_col >= 0:
                    curr_letter = self._col_index_to_letter(ddp_col)
                    prev_letter = self._col_index_to_letter(prev_ddp_col)
                    growth_letter = self._col_index_to_letter(growth_ddp_col)
                    formula = f'=IF(AND({curr_letter}{row_idx}<>"",{prev_letter}{row_idx}<>"",{prev_letter}{row_idx}<>0),({curr_letter}{row_idx}-{prev_letter}{row_idx})/{prev_letter}{row_idx},"")'
                    updates.append({
                        "range": f"{growth_letter}{row_idx}",
                        "values": [[formula]]
                    })

            if updates:
                # Batch update in chunks to avoid quota limits
                chunk_size = 500
                for i in range(0, len(updates), chunk_size):
                    chunk = updates[i:i + chunk_size]
                    ws.batch_update(chunk, value_input_option="USER_ENTERED")

                logger.info(f"Applied {len(updates)} formulas on {sheet_name}")
                self._log(spreadsheet_id, "ПРАЙС", "Формулы применены", f"Динамика цены: {len(updates)} ячеек", "✅")

        except Exception as e:
            logger.error(f"Apply price dynamics formulas failed: {e}", exc_info=True)

    async def _apply_price_calculation_formulas(self, spreadsheet_id: str, config: Dict):
        """
        Apply INDEX/MATCH formulas on 'Расчет цены' sheet.
        Matches GAS: Lib.updatePriceCalculationFormulas()

        Pulls data from 'Динамика цены' by ID-P lookup.
        """
        sheet_name = config.get("sheets", {}).get("price_calculation", "Расчет цены")
        dynamics_sheet = config.get("sheets", {}).get("price_dynamics", "Динамика цены")

        self._log(spreadsheet_id, "ПРАЙС", "Применение INDEX/MATCH", f"Лист: {sheet_name}", "📊")

        try:
            ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)
            values = ws.get_all_values()

            if not values or len(values) < 2:
                return

            headers = values[0]

            # Find ID-P column
            idp_col = self._find_column_index(headers, "ID-P", -1)
            if idp_col < 0:
                logger.warning("ID-P column not found in price calculation sheet")
                return

            idp_letter = self._col_index_to_letter(idp_col)

            # Columns to pull from Динамика цены via INDEX/MATCH
            # Format: (target_column_keyword, source_column_keyword)
            formula_mappings = [
                ("EXW  текущая", "EXW ALFASPA"),
                ("EXW  ALFASPA  текущая", "EXW ALFASPA"),
                ("Закупочная цена", "Закупочная цена"),
                ("DDP  -МОСКВА", "DDP-МОСКВА"),
            ]

            updates = []
            num_rows = len(values)

            for target_keyword, source_keyword in formula_mappings:
                target_col = self._find_column_index(headers, target_keyword, -1)
                if target_col < 0:
                    continue

                target_letter = self._col_index_to_letter(target_col)

                # INDEX/MATCH formula to pull data from Динамика цены
                # =IFERROR(INDEX('Динамика цены'!$A:$ZZ,MATCH($B2,'Динамика цены'!$B:$B,0),MATCH("source_keyword",'Динамика цены'!$1:$1,0)),"")
                for row_idx in range(2, num_rows + 1):
                    formula = f'=IFERROR(INDEX(\'{dynamics_sheet}\'!$A:$ZZ,MATCH({idp_letter}{row_idx},\'{dynamics_sheet}\'!$B:$B,0),MATCH("{source_keyword}",\'{dynamics_sheet}\'!$1:$1,0)),"")'
                    updates.append({
                        "range": f"{target_letter}{row_idx}",
                        "values": [[formula]]
                    })

            if updates:
                # Batch update in chunks
                chunk_size = 500
                for i in range(0, len(updates), chunk_size):
                    chunk = updates[i:i + chunk_size]
                    ws.batch_update(chunk, value_input_option="USER_ENTERED")

                logger.info(f"Applied {len(updates)} INDEX/MATCH formulas on {sheet_name}")
                self._log(spreadsheet_id, "ПРАЙС", "INDEX/MATCH применены", f"Расчет цены: {len(updates)} ячеек", "✅")

        except Exception as e:
            logger.error(f"Apply price calculation formulas failed: {e}", exc_info=True)

    async def _fill_idp_on_all_sheets_optimized(
        self,
        spreadsheet_id: str,
        config: Dict,
        id_to_idp: Dict[str, str]
    ):
        """
        Fill ID-P on ALL sheets.
        Each sheet is updated separately to avoid gspread range issues.

        Args:
            spreadsheet_id: Target spreadsheet
            config: Project configuration
            id_to_idp: Mapping of ID → ID-P values
        """
        if not id_to_idp:
            return

        # All sheets that need ID-P filling
        all_sheets = list(config.get("base_sheets_for_creation", []))
        primary = config.get("sheets", {}).get("primary", "Главная")
        if primary not in all_sheets:
            all_sheets.append(primary)

        self._log(
            spreadsheet_id,
            "ПРАЙС",
            "Заполнение ID-P",
            f"Листы: {', '.join(all_sheets)}",
            "🔄"
        )

        total_updates = 0

        for sheet_name in all_sheets:
            try:
                ws = self.sheets.get_worksheet(spreadsheet_id, sheet_name)
                values = ws.get_all_values()

                if not values:
                    continue

                headers = values[0]
                id_col = self._find_column_index(headers, "ID", -1)
                idp_col = self._find_column_index(headers, "ID-P", -1)

                if id_col < 0 or idp_col < 0:
                    logger.warning(f"ID or ID-P column not found in {sheet_name}")
                    continue

                idp_letter = self._col_index_to_letter(idp_col)

                # Collect updates for THIS sheet only (no sheet name in range)
                sheet_updates = []
                for row_idx, row in enumerate(values[1:], start=2):
                    id_val = self._as_trimmed_string(self._safe_get(row, id_col))
                    current_idp = self._as_trimmed_string(self._safe_get(row, idp_col))

                    if id_val and id_val in id_to_idp and not current_idp:
                        sheet_updates.append({
                            "range": f"{idp_letter}{row_idx}",
                            "values": [[id_to_idp[id_val]]]
                        })

                # Update this sheet
                if sheet_updates:
                    chunk_size = 500
                    for i in range(0, len(sheet_updates), chunk_size):
                        chunk = sheet_updates[i:i + chunk_size]
                        ws.batch_update(chunk, value_input_option="RAW")
                    total_updates += len(sheet_updates)
                    logger.info(f"Filled {len(sheet_updates)} ID-P values on {sheet_name}")

                # Small delay between sheets to avoid rate limits
                await asyncio.sleep(2)

            except Exception as e:
                logger.warning(f"Error processing sheet {sheet_name}: {e}")

        if total_updates > 0:
            self._log(
                spreadsheet_id,
                "ПРАЙС",
                "ID-P заполнен",
                f"Всего: {total_updates} ячеек на {len(all_sheets)} листах",
                "✅"
            )


# Singleton instance
_price_processor: Optional[PriceProcessor] = None


def get_price_processor(sheets_service: SheetsService, sync_service: Any, logging_service=None) -> PriceProcessor:
    """Get or create PriceProcessor instance"""
    global _price_processor
    if _price_processor is None:
        _price_processor = PriceProcessor(sheets_service, sync_service, logging_service)
    return _price_processor

# Export for modular routes
from src.services.sheets_service import sheets_service
from src.services.sync_service import sync_service
from src.services.logging_service import logging_service

price_processor = get_price_processor(sheets_service, sync_service, logging_service)
