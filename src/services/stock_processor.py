"""
Stock Processor Service

Handles loading of stock data from "-остатки" sheet to "Заказ" sheet.
Replaces GAS logic in 02Загрузка остатков.js.
"""

import re
import yaml
from pathlib import Path
from typing import List, Dict, Any, Optional
from src.utils.logger import logger
from src.services.sheets import SheetsService

class StockProcessor:
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
            logger.warning(f"Config file not found: {config_path}")
            return {}

        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)
            self._config_cache[project] = config
            return config
        except Exception as e:
            logger.error(f"Failed to load config for {project}: {e}")
            return {}

    async def process(self, spreadsheet_id: str, project: str, source_doc_id: Optional[str] = None) -> Dict[str, Any]:
        """
        Main processing method for stock loading.
        """
        self._log(spreadsheet_id, "ОСТАТКИ", f"Загрузка остатков ({project.upper()})", "Старт", "🚀")

        try:
            # 1. Load project config for sheet names
            config = self.load_config(project)
            
            # Resolve sheet names from config or use defaults
            order_sheet_name = config.get("sheets", {}).get("order_form", "Заказ")
            stocks_sheet_name = config.get("sheets", {}).get("stocks", "-остатки")
            
            # 2. Open worksheets
            order_ws = self.sheets.get_worksheet(spreadsheet_id, order_sheet_name)
            
            # Resolve source document ID
            # Priority: 1. source_doc_id argument, 2. source.doc_id from config, 3. current spreadsheet
            stocks_doc_id = source_doc_id or config.get("source", {}).get("doc_id") or spreadsheet_id
            
            # If config has specific sheet name for source stocks, use it
            stocks_source_sheet_name = config.get("source", {}).get("sheets", {}).get("stocks", stocks_sheet_name)
            
            self._log(spreadsheet_id, "ОСТАТКИ", "Открытие листов", f"Заказ: {order_sheet_name}, Источник ID: {stocks_doc_id[:10]}..., Лист: {stocks_source_sheet_name}")
            
            stocks_ws = self.sheets.get_worksheet(stocks_doc_id, stocks_source_sheet_name)
            
            # 3. Read Order sheet data
            order_values = order_ws.get_all_values()
            if not order_values:
                raise ValueError(f"Лист '{order_sheet_name}' пуст")
                
            # Find headers in Order sheet
            order_headers = order_values[0]
            
            # Map column indices using configuration if available
            mapping = config.get("column_mapping", {})
            
            order_idx = {
                "article": self._find_column_index(order_headers, mapping.get("article_rus", "Арт. Рус")),
                "sales": self._find_column_index(order_headers, "ПРОДАЖИ"),
                "writtenOff": self._find_column_index(order_headers, "СПИСАНО"),
                "stock": self._find_column_index(order_headers, "Остаток"),
                "inTransit": self._find_column_index(order_headers, "товар в ПУТИ"),
                "reserve": self._find_column_index(order_headers, "РЕЗЕРВ"),
                "qty1": self._find_column_index(order_headers, "Остаток 1"),
                "exp1": self._find_column_index(order_headers, "СГ 1"),
                "qty2": self._find_column_index(order_headers, "Остаток 2"),
                "exp2": self._find_column_index(order_headers, "СГ 2"),
                "qty3": self._find_column_index(order_headers, "Остаток3"),
                "exp3": self._find_column_index(order_headers, "СГ 3"),
            }
            if order_idx["qty3"] == -1:
                order_idx["qty3"] = self._find_column_index(order_headers, "Остаток 3")

            if order_idx["article"] == -1:
                raise ValueError(f"На листе '{order_sheet_name}' не найден столбец артикула")

            # 4. Read Stocks sheet data (Headers in row 3 or from config)
            # Fetch all values
            stocks_all_values = stocks_ws.get_all_values()
            
            # Determine header row (GAS logic often expects row 3)
            header_row_idx = 2 # Default 0-based index for row 3
            if len(stocks_all_values) <= header_row_idx:
                 # If sheet is small, try row 1
                 header_row_idx = 0
                 
            if len(stocks_all_values) <= header_row_idx:
                raise ValueError(f"На листе '{stocks_source_sheet_name}' недостаточно строк")
            
            stocks_headers = stocks_all_values[header_row_idx]
            stocks_data_rows = stocks_all_values[header_row_idx+1:]
            
            stocks_idx = {
                "article": self._find_column_index(stocks_headers, "Артикул"),
                "sold": self._find_column_index(stocks_headers, "Продано за период"),
                "writtenOff": self._find_column_index(stocks_headers, "Списано за период"),
                "stock": self._find_column_index(stocks_headers, "Остаток на текущий день"),
                "reserve": self._find_column_index(stocks_headers, "Из них в резерве"),
                "total": self._find_column_index(stocks_headers, "Всего"),
                "expiry": self._find_column_index(stocks_headers, "Срок годности"),
            }
            if stocks_idx["article"] == -1:
                stocks_idx["article"] = self._find_column_index(stocks_headers, "Арт. Рус")
            
            if stocks_idx["article"] == -1:
                # Try Col C (index 2) as fallback like in GAS
                if len(stocks_headers) > 2 and "артикул" in str(stocks_headers[2]).lower():
                    stocks_idx["article"] = 2
                else:
                    raise ValueError(f"На листе '{stocks_source_sheet_name}' не найден столбец Артикул")

            # 5. Build Stocks Map
            stocks_map = {}
            current_article = ""
            current_sold = None
            current_written = None
            current_stock = None
            
            for row in stocks_data_rows:
                # Article detection
                article_val = self._safe_get(row, stocks_idx["article"]).strip()
                if article_val:
                    current_article = article_val
                    current_sold = None
                    current_written = None
                    current_stock = None
                
                if not current_article:
                    continue
                
                # Fetch values
                if stocks_idx["sold"] != -1:
                    val = self._safe_get(row, stocks_idx["sold"])
                    if val != "": current_sold = self._parse_number(val)
                if stocks_idx["writtenOff"] != -1:
                    val = self._safe_get(row, stocks_idx["writtenOff"])
                    if val != "": current_written = self._parse_number(val)
                if stocks_idx["stock"] != -1:
                    val = self._safe_get(row, stocks_idx["stock"])
                    if val != "": current_stock = self._parse_number(val)
                
                if current_article not in stocks_map:
                    stocks_map[current_article] = {
                        "sales": current_sold,
                        "writtenOff": current_written,
                        "stock": current_stock,
                        "reserve": 0,
                        "batches": [] 
                    }
                else:
                    entry = stocks_map[current_article]
                    if current_sold is not None: entry["sales"] = current_sold
                    if current_written is not None: entry["writtenOff"] = current_written
                    if current_stock is not None: entry["stock"] = current_stock
                
                # Reserve
                if stocks_idx["reserve"] != -1:
                    reserve_val = self._parse_number(self._safe_get(row, stocks_idx["reserve"]))
                    if reserve_val:
                        stocks_map[current_article]["reserve"] += reserve_val
                
                # Batches
                if stocks_idx["total"] != -1:
                    total_val = self._parse_number(self._safe_get(row, stocks_idx["total"]))
                    if total_val:
                        expiry_val = self._safe_get(row, stocks_idx["expiry"])
                        stocks_map[current_article]["batches"].append({
                            "qty": total_val,
                            "expiry": expiry_val
                        })

            # 6. Prepare Updates for Order sheet
            num_rows = len(order_values) - 1
            if num_rows <= 0:
                return {"status": "success", "message": "Нет строк для обновления", "updated_rows": 0}

            target_cols = ["sales", "writtenOff", "stock", "inTransit", "reserve", "qty1", "exp1", "qty2", "exp2", "qty3", "exp3"]
            updates_by_col = {col_key: [[""]] * num_rows for col_key in target_cols}
            
            for i in range(1, len(order_values)):
                order_row = order_values[i]
                article = order_row[order_idx["article"]].strip()
                stats = stocks_map.get(article)
                
                row_idx = i - 1
                if stats:
                    if order_idx["sales"] != -1: updates_by_col["sales"][row_idx] = [stats["sales"] if stats["sales"] is not None else ""]
                    if order_idx["writtenOff"] != -1: updates_by_col["writtenOff"][row_idx] = [stats["writtenOff"] if stats["writtenOff"] is not None else ""]
                    if order_idx["stock"] != -1: updates_by_col["stock"][row_idx] = [stats["stock"] if stats["stock"] is not None else ""]
                    if order_idx["reserve"] != -1: updates_by_col["reserve"][row_idx] = [stats["reserve"] if stats["reserve"] > 0 else ""]
                    
                    batches = stats["batches"][:3]
                    for b_idx, batch in enumerate(batches):
                        q_key = f"qty{b_idx+1}"
                        e_key = f"exp{b_idx+1}"
                        updates_by_col[q_key][row_idx] = [batch["qty"]]
                        updates_by_col[e_key][row_idx] = [batch["expiry"]]

            # 7. Apply Batch Updates
            from gspread.utils import rowcol_to_a1
            batch_data = []
            
            for col_key, values in updates_by_col.items():
                col_idx = order_idx[col_key]
                if col_idx != -1:
                    range_a1 = f"{rowcol_to_a1(2, col_idx + 1)}:{rowcol_to_a1(len(order_values), col_idx + 1)}"
                    batch_data.append({
                        "range": range_a1,
                        "values": values
                    })
            
            if batch_data:
                order_ws.batch_update(batch_data, value_input_option='USER_ENTERED')
            
            self._log(spreadsheet_id, "ОСТАТКИ", "Загрузка остатков завершена", f"Обновлено строк: {num_rows}", "✅")
            
            return {
                "status": "success",
                "message": f"Загрузка остатков завершена. Обновлено строк: {num_rows}.",
                "updated_rows": num_rows,
                "source_doc": stocks_doc_id
            }

        except Exception as e:
            logger.error(f"Stock processing failed: {e}", exc_info=True)
            self._log(spreadsheet_id, "ОСТАТКИ", "Ошибка загрузки остатков", str(e), "❌")
            raise

    def _find_column_index(self, headers: List[str], target: str) -> int:
        norm_target = target.strip().lower()
        for i, h in enumerate(headers):
            if h.strip().lower() == norm_target:
                return i
        return -1

    def _safe_get(self, row: List[Any], idx: int) -> str:
        if idx < 0 or idx >= len(row):
            return ""
        return str(row[idx]) if row[idx] is not None else ""

    def _parse_number(self, val: Any) -> Optional[float]:
        if val is None or val == "":
            return 0.0
        try:
            s = str(val).replace(",", ".").replace("\xa0", "").replace(" ", "")
            return float(s)
        except:
            return 0.0
