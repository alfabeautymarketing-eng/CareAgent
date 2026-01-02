"""
Order stage filtering service.

Handles filtering the "Заказ" sheet by different stages:
- all: Show all data
- order: Show order-related columns
- promotions: Show promotion-related columns
- set: Show set-related columns
- price: Show price-related columns

Ported from GAS 06OrderForm.js (showAllOrderData, showOrderStage, etc.)
"""
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from enum import Enum

from src.utils.logger import logger


class OrderStageType(str, Enum):
    """Order stage types for filtering."""
    ALL = "all"
    ORDER = "order"
    PROMOTIONS = "promotions"
    SET = "set"
    PRICE = "price"


@dataclass
class StageFilterConfig:
    """Configuration for stage filtering."""
    columns_to_hide: List[str]
    hide_archived: bool = True
    hide_not_ordered: bool = False
    hide_categories: Optional[List[str]] = None


@dataclass
class FilterResult:
    """Result of filter operation."""
    stage: str
    visible_rows: int
    hidden_rows: int
    hidden_columns: int
    message: str


class OrderService:
    """
    Service for managing order sheet stages.

    Provides filtering by stage (order, promotions, set, price)
    by showing/hiding columns and rows.
    """

    # Columns to hide for each stage
    STAGE_CONFIGS = {
        OrderStageType.ORDER: StageFilterConfig(
            columns_to_hide=[
                "АКЦИИ",
                "условие#",
                "коментарий#",
                "Срок#",
                "Набор",
                "условие",
                "коментарий",
                "Срок",
                "Добавить в прайс",
                "Категория товара",
                "Группа линии",
                "Линия Прайс",
                "Примечание",
                "Продублировать"
            ],
            hide_archived=True,
            hide_not_ordered=False
        ),
        OrderStageType.PROMOTIONS: StageFilterConfig(
            columns_to_hide=[
                "Необходимо заказать, шт",
                "Предварительный заказ",
                "заказ в упаковках",
                "ЗАКАЗ",
                "EXW  ALFASPA  текущая, €",
                "Предварительная сумма заказа",
                "Набор",
                "условие",
                "коментарий",
                "Срок",
                "Добавить в прайс",
                "Категория товара",
                "Группа линии",
                "Линия Прайс",
                "Примечание",
                "Продублировать"
            ],
            hide_archived=True,
            hide_not_ordered=True
        ),
        OrderStageType.SET: StageFilterConfig(
            columns_to_hide=[
                "Необходимо заказать, шт",
                "Предварительный заказ",
                "заказ в упаковках",
                "ЗАКАЗ",
                "EXW  ALFASPA  текущая, €",
                "Предварительная сумма заказа",
                "АКЦИИ",
                "условие#",
                "коментарий#",
                "Срок#",
                "Добавить в прайс",
                "Категория товара",
                "Группа линии",
                "Линия Прайс",
                "Примечание",
                "Продублировать"
            ],
            hide_archived=True,
            hide_not_ordered=True
        ),
        OrderStageType.PRICE: StageFilterConfig(
            columns_to_hide=[
                "Остаток 1",
                "СГ 1",
                "Остаток 2",
                "СГ 2",
                "Остаток3",
                "СГ 3",
                "СПИСАНО",
                "РЕЗЕРВ",
                "Среднемесячные продажи, шт",
                "Потребность на 6 месяцев, шт",
                "На сколько мес запас",
                "Необходимо заказать, шт",
                "Предварительный заказ",
                "заказ в упаковках",
                "ЗАКАЗ",
                "EXW  ALFASPA  текущая, €",
                "Предварительная сумма заказа",
                "АКЦИИ",
                "условие#",
                "коментарий#",
                "Срок#",
                "Набор",
                "условие",
                "коментарий",
                "Срок"
            ],
            hide_archived=True,
            hide_not_ordered=True,
            hide_categories=["Пробники", "Mini Size", "Тестер"]
        )
    }

    # Statuses that indicate archived products
    ARCHIVED_STATUSES = ["Снят архив", "Архив"]

    # Statuses that indicate "not ordered" products
    NOT_ORDERED_STATUSES = ["Не заказываем"]

    def __init__(self, sheets_service=None):
        """
        Initialize order service.

        Args:
            sheets_service: Google Sheets service for reading/writing
        """
        self.sheets = sheets_service

    def get_stage_config(self, stage: OrderStageType) -> Optional[StageFilterConfig]:
        """Get configuration for a specific stage."""
        return self.STAGE_CONFIGS.get(stage)

    def get_column_indexes(
        self,
        headers: List[str],
        column_names: List[str]
    ) -> Dict[str, int]:
        """
        Get column indexes for given column names.

        Args:
            headers: List of header names from the sheet
            column_names: List of column names to find

        Returns:
            Dictionary mapping column name to 1-based index
        """
        indexes = {}

        for col_name in column_names:
            # Try exact match first
            if col_name in headers:
                indexes[col_name] = headers.index(col_name) + 1
                continue

            # Try partial match (starts with)
            for i, header in enumerate(headers):
                header_str = str(header or "")
                if header_str.startswith(col_name):
                    indexes[col_name] = i + 1
                    break

        return indexes

    def filter_rows_by_status(
        self,
        data: List[List[Any]],
        headers: List[str],
        config: StageFilterConfig
    ) -> List[int]:
        """
        Determine which rows should be hidden based on filter config.

        Args:
            data: 2D array of row data (excluding header)
            headers: List of header names
            config: Filter configuration

        Returns:
            List of row numbers (1-based, data rows) to hide
        """
        rows_to_hide = []

        # Find column indexes
        status_idx = headers.index("Статус") if "Статус" in headers else -1
        remainder_idx = headers.index("Остаток") if "Остаток" in headers else -1
        sales_idx = headers.index("ПРОДАЖИ") if "ПРОДАЖИ" in headers else -1
        category_idx = headers.index("Категория товара") if "Категория товара" in headers else -1

        if status_idx == -1:
            logger.warning("Column 'Статус' not found")
            return []

        for i, row in enumerate(data):
            row_num = i + 2  # 1-based, skip header

            status = str(row[status_idx] if status_idx < len(row) else "").strip()
            remainder = row[remainder_idx] if remainder_idx >= 0 and remainder_idx < len(row) else ""
            sales = row[sales_idx] if sales_idx >= 0 and sales_idx < len(row) else ""

            should_hide = False

            # Check archived status
            if config.hide_archived:
                if status in self.ARCHIVED_STATUSES:
                    # For archived, only hide if both remainder and sales are empty
                    if not remainder and not sales:
                        should_hide = True

            # Check not ordered status
            if config.hide_not_ordered:
                if status in self.NOT_ORDERED_STATUSES:
                    should_hide = True

            # Check categories to hide
            if config.hide_categories and category_idx >= 0:
                category = str(row[category_idx] if category_idx < len(row) else "").strip()
                if category in config.hide_categories:
                    should_hide = True

            if should_hide:
                rows_to_hide.append(row_num)

        return rows_to_hide

    async def apply_stage_filter(
        self,
        spreadsheet_id: str,
        stage: OrderStageType,
        sheet_name: str = "Заказ",
        dry_run: bool = False
    ) -> FilterResult:
        """
        Apply stage filter to the order sheet.

        Args:
            spreadsheet_id: ID of the spreadsheet
            stage: Stage type to apply
            sheet_name: Name of the order sheet
            dry_run: If True, don't actually apply changes

        Returns:
            FilterResult with statistics
        """
        if not self.sheets:
            raise ValueError("Sheets service not initialized")

        try:
            logger.info(f"Applying stage filter: {stage.value}")

            # Get sheet metadata
            metadata = await self.sheets.get_sheet_metadata(spreadsheet_id, sheet_name)
            sheet_id = metadata.get("sheetId", 0)
            max_rows = metadata.get("gridProperties", {}).get("rowCount", 1000)
            max_cols = metadata.get("gridProperties", {}).get("columnCount", 50)

            # Read headers
            headers_result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{sheet_name}'!1:1"
            )
            headers = headers_result.get("values", [[]])[0]

            if stage == OrderStageType.ALL:
                # Show all columns and rows
                if not dry_run:
                    await self._show_all(spreadsheet_id, sheet_id, max_rows, max_cols)

                return FilterResult(
                    stage=stage.value,
                    visible_rows=max_rows - 1,
                    hidden_rows=0,
                    hidden_columns=0,
                    message="All filters removed"
                )

            # Get stage configuration
            config = self.get_stage_config(stage)
            if not config:
                raise ValueError(f"Unknown stage: {stage.value}")

            # Find columns to hide
            col_indexes = self.get_column_indexes(headers, config.columns_to_hide)
            columns_to_hide = list(col_indexes.values())

            # Read all data for row filtering
            data_result = await self.sheets.get(
                spreadsheet_id=spreadsheet_id,
                range=f"'{sheet_name}'!A2:ZZ"
            )
            data = data_result.get("values", [])

            # Determine rows to hide
            rows_to_hide = self.filter_rows_by_status(data, headers, config)

            if not dry_run:
                # First show all
                await self._show_all(spreadsheet_id, sheet_id, max_rows, max_cols)

                # Then hide columns
                if columns_to_hide:
                    await self._hide_columns(spreadsheet_id, sheet_id, columns_to_hide)

                # Then hide rows
                if rows_to_hide:
                    await self._hide_rows(spreadsheet_id, sheet_id, rows_to_hide)

            visible_rows = len(data) - len(rows_to_hide)

            return FilterResult(
                stage=stage.value,
                visible_rows=visible_rows,
                hidden_rows=len(rows_to_hide),
                hidden_columns=len(columns_to_hide),
                message=f"Filter '{stage.value}' applied"
            )

        except Exception as e:
            logger.error(f"Stage filter error: {e}")
            raise

    async def _show_all(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        max_rows: int,
        max_cols: int
    ):
        """Show all rows and columns."""
        requests = [
            # Show all columns
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": max_cols
                    },
                    "properties": {"hiddenByUser": False},
                    "fields": "hiddenByUser"
                }
            },
            # Show all rows (except header)
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": 1,
                        "endIndex": max_rows
                    },
                    "properties": {"hiddenByUser": False},
                    "fields": "hiddenByUser"
                }
            }
        ]

        await self.sheets.batch_update_spreadsheet(spreadsheet_id, requests)

    async def _hide_columns(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        column_indexes: List[int]
    ):
        """Hide specified columns (1-based indexes)."""
        requests = []

        for col_idx in column_indexes:
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": col_idx - 1,  # Convert to 0-based
                        "endIndex": col_idx
                    },
                    "properties": {"hiddenByUser": True},
                    "fields": "hiddenByUser"
                }
            })

        if requests:
            await self.sheets.batch_update_spreadsheet(spreadsheet_id, requests)
            logger.info(f"Hidden {len(requests)} columns")

    async def _hide_rows(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        row_numbers: List[int]
    ):
        """Hide specified rows (1-based row numbers)."""
        if not row_numbers:
            return

        # Optimize by grouping consecutive rows
        requests = []
        sorted_rows = sorted(row_numbers)

        i = 0
        while i < len(sorted_rows):
            start = sorted_rows[i] - 1  # Convert to 0-based
            end = start + 1

            # Find consecutive rows
            while i + 1 < len(sorted_rows) and sorted_rows[i + 1] == sorted_rows[i] + 1:
                i += 1
                end += 1

            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start,
                        "endIndex": end
                    },
                    "properties": {"hiddenByUser": True},
                    "fields": "hiddenByUser"
                }
            })
            i += 1

        if requests:
            await self.sheets.batch_update_spreadsheet(spreadsheet_id, requests)
            logger.info(f"Hidden {len(row_numbers)} rows in {len(requests)} requests")


# Singleton instance
_order_service: Optional[OrderService] = None


def get_order_service(sheets_service=None) -> OrderService:
    """Get or create OrderService singleton."""
    global _order_service
    if _order_service is None:
        _order_service = OrderService(sheets_service)
    return _order_service
