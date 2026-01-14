import time
import os
import yaml
import asyncio
from pathlib import Path
from typing import List, Dict, Optional, Any, Tuple
from pydantic import BaseModel
from src.services.sheets import SheetsService
from src.services.function_log_service import FunctionLogService, FunctionLogContext
from src.utils.logger import logger
import gspread
from gspread.utils import rowcol_to_a1

class SyncRule(BaseModel):
    """Модель правила синхронизации с поддержкой bidirectional режима."""
    id: str
    enabled: bool
    category: str

    # Режим работы: "unidirectional" (одностороннее) или "bidirectional" (двустороннее)
    mode: str = "unidirectional"

    # Для unidirectional (обратная совместимость)
    source_sheet: str = ""
    source_header: str = ""
    target_sheet: str = ""
    target_header: str = ""

    # Для bidirectional: оба листа равноправны
    sheet_a: Optional[str] = None
    header_a: Optional[str] = None
    sheet_b: Optional[str] = None
    header_b: Optional[str] = None

    is_external: bool = False
    target_doc_id: Optional[str] = None

class SyncService:
    def __init__(
        self,
        logging_service: Optional[Any] = None,
        sync_log_service: Optional[Any] = None,
        function_log_service: Optional[FunctionLogService] = None,
    ):
        self.sheets_service = SheetsService()
        self.logging_service = logging_service
        self.sync_log_service = sync_log_service
        self.function_log_service = function_log_service or FunctionLogService()
        self._rules_cache: List[SyncRule] = []
        self._rules_cache_time = 0
        self._rules_cache_ttl = 300  # 5 minutes

        # Защита от циклов: блокировки ячеек с TTL
        self._sync_locks: Dict[str, float] = {}
        self._sync_lock_ttl = 3.0  # секунд

        self.sheet_codes = {
            "Главная": "Г",
            "Сертификация": "С",
            "Этикетки": "Э",
            "Заказ": "З",
            "Динамика цены": "Д",
            "Расчет цены": "Р",
            "Прайс": "П",
            "ABC-Анализ": "А",
            "Журнал синхро": "J",
            "Логи": "L",
        }

    # === Защита от циклов ===

    def _is_cell_locked(self, spreadsheet_id: str, row_key: str, header: str) -> bool:
        """Проверяет, заблокирована ли ячейка от повторной синхронизации."""
        key = f"{spreadsheet_id}:{row_key}:{header}"
        now = time.time()

        # Очистка устаревших блокировок
        self._sync_locks = {
            k: v for k, v in self._sync_locks.items()
            if now - v < self._sync_lock_ttl
        }

        return key in self._sync_locks

    def _lock_cell(self, spreadsheet_id: str, row_key: str, header: str):
        """Блокирует ячейку на время TTL для предотвращения циклов."""
        key = f"{spreadsheet_id}:{row_key}:{header}"
        self._sync_locks[key] = time.time()

    def _get_matching_rules(self, rules: List[SyncRule], sheet_name: str, header: str) -> List[SyncRule]:
        """Находит правила для указанного листа и заголовка (с учетом bidirectional)."""
        matching = []
        for rule in rules:
            if not rule.enabled:
                continue

            if rule.mode == "bidirectional":
                # Bidirectional: срабатывает при изменении на любом из двух листов
                if (rule.sheet_a == sheet_name and rule.header_a == header) or \
                   (rule.sheet_b == sheet_name and rule.header_b == header):
                    matching.append(rule)
            else:
                # Unidirectional: только source -> target
                if rule.source_sheet == sheet_name and rule.source_header == header:
                    matching.append(rule)

        return matching

    def _get_target_for_rule(self, rule: SyncRule, source_sheet: str) -> Tuple[str, str]:
        """Определяет целевой лист и заголовок для правила."""
        if rule.mode == "bidirectional":
            if source_sheet == rule.sheet_a:
                return rule.sheet_b or "", rule.header_b or ""
            else:
                return rule.sheet_a or "", rule.header_a or ""
        else:
            return rule.target_sheet, rule.target_header

    def _make_rule_id(self, idx: int, source_sheet: str, target_sheet: str, header: str) -> str:
        src_code = self.sheet_codes.get(source_sheet, (source_sheet[:1] or "?"))
        tgt_code = self.sheet_codes.get(target_sheet, (target_sheet[:1] or "?"))
        return f"{idx:03d}-{src_code}-{tgt_code}({header})"

    def _save_rules_yaml(self, spreadsheet_id: str, rules: List[SyncRule]) -> None:
        path = Path("config") / "rules" / f"{spreadsheet_id}.yaml"
        path.parent.mkdir(parents=True, exist_ok=True)
        data = [r.model_dump() for r in rules]
        path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")

    def list_rules(
        self,
        spreadsheet_id: str,
        force_reload: bool = False,
        include_disabled: bool = True
    ) -> List[Dict[str, Any]]:
        if include_disabled:
            if force_reload:
                self._rules_cache = []
                self._rules_cache_time = 0
            rules = self._load_rules_raw(spreadsheet_id)
        else:
            rules = self._load_rules(spreadsheet_id, force_reload=force_reload)
        return [r.model_dump() for r in rules]

    def save_rules(self, spreadsheet_id: str, rules_payload: List[Dict[str, Any]], validate_headers: bool = True) -> List[Dict[str, Any]]:
        headers_cache: Dict[Tuple[str, str], List[str]] = {}
        normalized: List[SyncRule] = []

        def get_headers(doc_id: str, sheet_name: str) -> List[str]:
            key = (doc_id, sheet_name)
            if key not in headers_cache:
                ws = self.sheets_service.get_worksheet(doc_id, sheet_name)
                headers_cache[key] = [str(h or "").strip() for h in ws.row_values(1)]
            return headers_cache[key]

        for idx, r in enumerate(rules_payload, start=1):
            mode = str(r.get("mode", "unidirectional")).strip() or "unidirectional"
            category = str(r.get("category", "")).strip()
            enabled = bool(r.get("enabled", True))
            is_external = bool(r.get("is_external", False))
            target_doc_id = str(r.get("target_doc_id", "")).strip() or None

            if mode == "bidirectional":
                if is_external:
                    raise ValueError(f"Bidirectional правило (#{idx}) не поддерживает внешний документ")

                sheet_a = str(r.get("sheet_a", "")).strip()
                header_a = str(r.get("header_a", "")).strip()
                sheet_b = str(r.get("sheet_b", "")).strip()
                header_b = str(r.get("header_b", "")).strip()

                if not sheet_a or not header_a or not sheet_b or not header_b:
                    raise ValueError(f"Неполное bidirectional правило (#{idx}): нужны sheet_a, header_a, sheet_b, header_b")

                if validate_headers:
                    headers_a = get_headers(spreadsheet_id, sheet_a)
                    headers_b = get_headers(spreadsheet_id, sheet_b)
                    if header_a not in headers_a:
                        raise ValueError(f"Столбец '{header_a}' не найден на листе '{sheet_a}'")
                    if header_b not in headers_b:
                        raise ValueError(f"Столбец '{header_b}' не найден на листе '{sheet_b}'")

                rule_id = self._make_bidirectional_id(idx, sheet_a, sheet_b, header_a)
                normalized.append(
                    SyncRule(
                        id=rule_id,
                        enabled=enabled,
                        category=category,
                        mode=mode,
                        source_sheet="",
                        source_header="",
                        target_sheet="",
                        target_header="",
                        sheet_a=sheet_a,
                        header_a=header_a,
                        sheet_b=sheet_b,
                        header_b=header_b,
                        is_external=False,
                        target_doc_id=None,
                    )
                )
                continue

            if mode != "unidirectional":
                raise ValueError(f"Неизвестный режим rule (#{idx}): {mode}")

            src_sheet = str(r.get("source_sheet", "")).strip()
            tgt_sheet = str(r.get("target_sheet", "")).strip()
            src_header = str(r.get("source_header", "")).strip()
            tgt_header = str(r.get("target_header", "")).strip()

            if not src_sheet or not src_header or not tgt_sheet or not tgt_header:
                raise ValueError(f"Неполное правило (#{idx}): нужны source_sheet, source_header, target_sheet, target_header")

            if is_external and not target_doc_id:
                raise ValueError(f"Для внешнего правила (#{idx}) нужен target_doc_id")

            if validate_headers:
                target_doc = target_doc_id if is_external else spreadsheet_id
                headers = get_headers(target_doc, tgt_sheet)
                if tgt_header not in headers:
                    raise ValueError(f"Столбец '{tgt_header}' не найден на листе '{tgt_sheet}'")

            rule_id = self._make_rule_id(idx, src_sheet, tgt_sheet, src_header)
            normalized.append(
                SyncRule(
                    id=rule_id,
                    enabled=enabled,
                    category=category,
                    mode=mode,
                    source_sheet=src_sheet,
                    source_header=src_header,
                    target_sheet=tgt_sheet,
                    target_header=tgt_header,
                    sheet_a=None,
                    header_a=None,
                    sheet_b=None,
                    header_b=None,
                    is_external=is_external,
                    target_doc_id=target_doc_id,
                )
            )

        self._save_rules_yaml(spreadsheet_id, normalized)
        self._rules_cache = [r for r in normalized if r.enabled]
        self._rules_cache_time = time.time()
        logger.info("sync_rules_saved", spreadsheet_id=spreadsheet_id, count=len(normalized))
        return [r.model_dump() for r in normalized]

    def _get_project_prefix(self, project_key: str) -> str:
        """Get standard brand prefix for project."""
        prefixes = {
            "MT": "MT-",
            "SS": "SS-",
            "SK": "SK-"
        }
        return prefixes.get(project_key.upper(), "ID-")

    def _get_next_id(self, spreadsheet_id: str, prefix: str) -> str:
        """Find max numeric ID with prefix in 'Главная' and return next one."""
        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, "Главная")
            col_a = ws.col_values(1)
            
            max_num = 0
            for val in col_a:
                val = str(val).strip()
                if val.startswith(prefix):
                    try:
                        num_part = val[len(prefix):]
                        if num_part.isdigit():
                            num = int(num_part)
                            if num > max_num:
                                max_num = num
                    except (ValueError, TypeError):
                        continue
            
            next_num = max_num + 1
            return f"{prefix}{next_num:03d}"
        except Exception as e:
            logger.error("get_next_id_failed", spreadsheet_id=spreadsheet_id, prefix=prefix, error=str(e))
            # Fallback to a timestamp based unique ID if calculation fails
            return f"{prefix}NEW-{int(time.time())}"


    def _load_rules(self, spreadsheet_id: str, force_reload: bool = False) -> List[SyncRule]:
        """
        Load sync rules from YAML file on server.
        NOTE: Fallback to Google Sheet "Правила синхро" has been removed.
        All rules must be stored in config/rules/<spreadsheet_id>.yaml
        """
        if not force_reload and (time.time() - self._rules_cache_time < self._rules_cache_ttl) and self._rules_cache:
            return self._rules_cache

        try:
            rules_path = os.path.join("config", "rules", f"{spreadsheet_id}.yaml")
            if not os.path.exists(rules_path):
                logger.warning("rules_file_not_found", path=rules_path)
                return []

            with open(rules_path, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or []

            rules = []
            for row in data:
                try:
                    mode = str(row.get("mode", "unidirectional")).strip() or "unidirectional"

                    rule = SyncRule(
                        id=str(row.get("id", "")).strip(),
                        enabled=bool(row.get("enabled", True)),
                        category=str(row.get("category", "")).strip(),
                        mode=mode,
                        # Unidirectional fields
                        source_sheet=str(row.get("source_sheet", "")).strip(),
                        source_header=str(row.get("source_header", "")).strip(),
                        target_sheet=str(row.get("target_sheet", "")).strip(),
                        target_header=str(row.get("target_header", "")).strip(),
                        # Bidirectional fields
                        sheet_a=str(row.get("sheet_a", "")).strip() or None,
                        header_a=str(row.get("header_a", "")).strip() or None,
                        sheet_b=str(row.get("sheet_b", "")).strip() or None,
                        header_b=str(row.get("header_b", "")).strip() or None,
                        is_external=bool(row.get("is_external", False)),
                        target_doc_id=(str(row.get("target_doc_id", "")).strip() or None)
                    )

                    # Validation
                    if mode == "bidirectional":
                        if not rule.sheet_a or not rule.header_a or not rule.sheet_b or not rule.header_b:
                            continue
                        if rule.is_external:
                            continue
                    else:
                        if not rule.source_sheet or not rule.source_header or not rule.target_sheet or not rule.target_header:
                            continue

                    if rule.is_external and not rule.target_doc_id:
                        continue
                    if not rule.enabled:
                        continue

                    rules.append(rule)
                except Exception:
                    continue

            self._rules_cache = rules
            self._rules_cache_time = time.time()
            logger.info("sync_rules_loaded_from_yaml", count=len(rules), source=rules_path)
            return rules

        except Exception as e:
            logger.error("load_rules_failed", error=str(e))
            return self._rules_cache

    def sync_row(self, spreadsheet_id: str, sheet_name: str, row_number: int, row_key: str, project: str) -> Dict[str, Any]:
        """
        Sync a single row to target sheets based on rules.
        """
        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "СИНХРО", f"Синхронизация строки {sheet_name}#{row_number}", f"ID={row_key}, Project={project}", "🔄 ЗАПУСК")
            
        rules = self._load_rules(spreadsheet_id)
        effective_rules = [
            r for r in rules
            if r.enabled and (
                (r.mode == "bidirectional" and (r.sheet_a == sheet_name or r.sheet_b == sheet_name)) or
                (r.mode != "bidirectional" and r.source_sheet == sheet_name)
            )
        ]
        
        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "СИНХРО", "Проверка правил", f"Найдено {len(effective_rules)} активных правил для листа {sheet_name}", "ℹ️ ИНФО")

        results = []
        
        # 2. Get Source Data
        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
            
            # Fetch entire row data (mapped by header)
            row_values = ws.row_values(row_number)
            headers = ws.row_values(1)
            
            # Map headers to values
            row_data = {}
            for i, h in enumerate(headers):
                if i < len(row_values):
                    row_data[h] = row_values[i]
                    
        except Exception as e:
            logger.error("sync_row_fetch_failed", error=str(e))
            if self.logging_service:
                self.logging_service.add_log(spreadsheet_id, "СИНХРО", "Ошибка получения данных", f"Не удалось получить данные для строки {row_key} из листа {sheet_name}: {str(e)}", "❌ ОШИБКА")
            raise

        # 3. Apply Rules
        for rule in effective_rules:
            if rule.mode == "bidirectional":
                source_header = rule.header_a if rule.sheet_a == sheet_name else rule.header_b
            else:
                source_header = rule.source_header

            val = row_data.get(source_header or "")

            if not source_header or source_header not in row_data:
                logger.warning("source_header_missing", header=source_header)
                if self.logging_service:
                    self.logging_service.add_log(spreadsheet_id, "СИНХРО", "Пропуск правила", f"Исходный заголовок '{source_header}' не найден в строке {row_key} листа {sheet_name}", "⚠️ ПРЕДУПРЕЖДЕНИЕ")
                results.append({"rule_id": rule.id, "status": "skipped", "reason": "Source header missing"})
                continue

            res = self._apply_rule_extended(spreadsheet_id, rule, row_key, val, sheet_name, project=project)
            results.append(res)
            
        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "СИНХРО", "Синхронизация строки завершена", f"Обработано {len(effective_rules)} правил для ID={row_key}. Успешно: {len([r for r in results if r.get('status') == 'success'])}", "✅ УСПЕХ")
        return {"status": "success", "results": results}

    async def sync_row_async(self, spreadsheet_id: str, sheet_name: str, row_number: int, row_key: str, project: str) -> Dict[str, Any]:
        """
        Async version of sync_row with parallel rule application.
        Applies multiple rules concurrently using asyncio.gather().
        Performance: 3-10x faster for rows with multiple rules.
        """
        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "СИНХРО", f"[ASYNC] Синхронизация строки {sheet_name}#{row_number}", f"ID={row_key}, Project={project}", "🔄 ЗАПУСК")

        rules = self._load_rules(spreadsheet_id)
        effective_rules = [
            r for r in rules
            if r.enabled and (
                (r.mode == "bidirectional" and (r.sheet_a == sheet_name or r.sheet_b == sheet_name)) or
                (r.mode != "bidirectional" and r.source_sheet == sheet_name)
            )
        ]

        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "СИНХРО", "[ASYNC] Проверка правил", f"Найдено {len(effective_rules)} активных правил для листа {sheet_name}", "ℹ️ ИНФО")

        # Get source data
        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
            row_values = ws.row_values(row_number)
            headers = ws.row_values(1)

            row_data = {}
            for i, h in enumerate(headers):
                if i < len(row_values):
                    row_data[h] = row_values[i]

        except Exception as e:
            logger.error("sync_row_async_fetch_failed", error=str(e))
            if self.logging_service:
                self.logging_service.add_log(spreadsheet_id, "СИНХРО", "[ASYNC] Ошибка получения данных", f"Не удалось получить данные для строки {row_key}: {str(e)}", "❌ ОШИБКА")
            raise

        # Apply rules in parallel using asyncio.gather()
        tasks = []
        for rule in effective_rules:
            if rule.mode == "bidirectional":
                source_header = rule.header_a if rule.sheet_a == sheet_name else rule.header_b
            else:
                source_header = rule.source_header

            val = row_data.get(source_header or "")

            if not source_header or source_header not in row_data:
                logger.warning("source_header_missing", header=source_header)
                if self.logging_service:
                    self.logging_service.add_log(spreadsheet_id, "СИНХРО", "[ASYNC] Пропуск правила", f"Исходный заголовок '{source_header}' не найден", "⚠️ ПРЕДУПРЕЖДЕНИЕ")
                continue

            # Create coroutine for parallel execution
            task = asyncio.to_thread(self._apply_rule_extended, spreadsheet_id, rule, row_key, val, sheet_name, project)
            tasks.append(task)

        # Execute all rules in parallel
        results = []
        if tasks:
            results = await asyncio.gather(*tasks, return_exceptions=True)
            # Convert exceptions to result dicts
            results = [
                {"rule_id": getattr(task, 'rule_id', 'unknown'), "status": "error", "error": str(r)} if isinstance(r, Exception) else r
                for task, r in zip(effective_rules, results)
            ]

        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "СИНХРО", "[ASYNC] Синхронизация строки завершена", f"Обработано {len(effective_rules)} правил параллельно. Успешно: {len([r for r in results if r.get('status') == 'success'])}", "✅ УСПЕХ")
        return {"status": "success", "results": results, "async": True}

    def sync_full(self, spreadsheet_id: str, project: str, source_sheet: str) -> Dict[str, Any]:
        """
        Sync ALL rows in a sheet.
        """
        logger.info("sync_full_start", project=project, sheet=source_sheet)
        
        rules = self._load_rules(spreadsheet_id)
        matching_rules = [
            r for r in rules
            if (r.mode == "bidirectional" and (r.sheet_a == source_sheet or r.sheet_b == source_sheet)) or
               (r.mode != "bidirectional" and r.source_sheet == source_sheet)
        ]
        
        if not matching_rules:
            return {"status": "skipped", "reason": "No rules for this sheet"}
            
        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, source_sheet)
            all_values = ws.get_all_values()
            
            if len(all_values) < 2:
                return {"status": "success", "rows_processed": 0}
                
            headers = all_values[0]
            data_rows = all_values[1:]
            
            # Map header name to index
            header_map = {h: i for i, h in enumerate(headers)}
            
            # Verify ID column exists (Column A / Index 0)
            # In GAS logic, ID is usually assumed first column or configured?
            # We assume Col A is key.
            
            success_count = 0
            errors = []
            
            for row in data_rows:
                if not row: continue
                article = row[0] # Col A
                if not article: continue
                
                for rule in matching_rules:
                    if rule.mode == "bidirectional":
                        source_header = rule.header_a if rule.sheet_a == source_sheet else rule.header_b
                    else:
                        source_header = rule.source_header

                    src_idx = header_map.get(source_header or "")
                    if src_idx is not None and src_idx < len(row):
                        val = row[src_idx]

                        try:
                            self._apply_rule_extended(spreadsheet_id, rule, article, val, source_sheet, project=project)
                        except Exception as e:
                            errors.append(f"{article}-{rule.id}: {str(e)}")
                    else:
                        pass
                
                success_count += 1
                
            return {
                "status": "success", 
                "rows_processed": success_count, 
                "rules_count": len(matching_rules),
                "errors": errors[:10] # Return first 10 errors
            }
            
        except Exception as e:
            logger.error("sync_full_failed", error=str(e))
            raise

    async def sync_full_async(self, spreadsheet_id: str, project: str, source_sheet: str) -> Dict[str, Any]:
        """
        Async version of sync_full with parallel row processing.
        Processes multiple rows concurrently using asyncio.gather().
        Performance: 5-10x faster for full sheet syncs.
        """
        logger.info("sync_full_async_start", project=project, sheet=source_sheet)

        rules = self._load_rules(spreadsheet_id)
        matching_rules = [
            r for r in rules
            if (r.mode == "bidirectional" and (r.sheet_a == source_sheet or r.sheet_b == source_sheet)) or
               (r.mode != "bidirectional" and r.source_sheet == source_sheet)
        ]

        if not matching_rules:
            return {"status": "skipped", "reason": "No rules for this sheet"}

        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, source_sheet)
            all_values = ws.get_all_values()

            if len(all_values) < 2:
                return {"status": "success", "rows_processed": 0}

            headers = all_values[0]
            data_rows = all_values[1:]

            # Map header name to index
            header_map = {h: i for i, h in enumerate(headers)}

            success_count = 0
            errors = []

            # Create tasks for parallel row processing
            tasks = []
            row_indices = []

            for row_idx, row in enumerate(data_rows):
                if not row:
                    continue

                article = row[0]  # Col A
                if not article:
                    continue

                row_indices.append(article)

                # Create tasks for all rules for this row
                for rule in matching_rules:
                    if rule.mode == "bidirectional":
                        source_header = rule.header_a if rule.sheet_a == source_sheet else rule.header_b
                    else:
                        source_header = rule.source_header

                    src_idx = header_map.get(source_header or "")
                    if src_idx is not None and src_idx < len(row):
                        val = row[src_idx]
                        # Use to_thread to run blocking I/O in parallel
                        task = asyncio.to_thread(
                            self._apply_rule_extended,
                            spreadsheet_id, rule, article, val, source_sheet,
                            project=project
                        )
                        tasks.append(task)

                success_count += 1

            # Execute all tasks in parallel
            if tasks:
                results = await asyncio.gather(*tasks, return_exceptions=True)
                for result in results:
                    if isinstance(result, Exception):
                        errors.append(str(result))

            return {
                "status": "success",
                "rows_processed": success_count,
                "rules_count": len(matching_rules),
                "errors": errors[:10],
                "async": True
            }

        except Exception as e:
            logger.error("sync_full_async_failed", error=str(e))
            raise

    async def sync_event(self, spreadsheet_id: str, event_data: Dict[str, Any]):
        """
        Process a single onEdit event with cycle protection.
        Includes detailed function logging and summary logging.
        """
        sheet_name = event_data.get("sheet_name")
        row_idx = event_data.get("row")
        col_idx = event_data.get("col")
        source_header = event_data.get("header_name")
        new_value = event_data.get("value")
        row_key = event_data.get("row_key")
        project = event_data.get("project")

        # === ЗАЩИТА ОТ ЦИКЛОВ: Проверка sync_origin ===
        sync_origin = event_data.get("sync_origin", "user")
        if sync_origin == "sync":
            logger.info("skipping_sync_originated_event",
                       sheet=sheet_name, row_key=row_key, header=source_header)
            return {"status": "skipped", "reason": "Sync-originated change ignored"}

        # === ЗАЩИТА ОТ ЦИКЛОВ: Проверка блокировки ячейки ===
        if row_key and source_header and self._is_cell_locked(spreadsheet_id, row_key, source_header):
            logger.info("skipping_locked_cell",
                       sheet=sheet_name, row_key=row_key, header=source_header)
            return {"status": "skipped", "reason": "Cell locked (anti-cycle protection)"}

        # Пропускаем события на лог-листах (Логи удалены из системы)
        if sheet_name in {"Журнал синхро", "Logs"}:
            return {"status": "skipped", "reason": "Log sheet ignored"}

        self._log_to_session(
            spreadsheet_id, 
            "СИНХРОНИЗАЦИЯ", 
            f"Обработка события: лист '{sheet_name}', строка {row_idx}",
            f"Колонка: '{source_header}', Ключ: '{row_key}', Нов.значение: '{new_value}'",
            "⚙️ ПРОЦЕСС"
        )
        
        # 1. Load Rules
        rules = self._load_rules(spreadsheet_id)
        
        # 2. Identify Source Header
        # We need to know the header of the column that was edited.
        # This requires reading the header row (usually row 1) of the edited sheet.
        # Optimization: Pass header from GAS? Or fetch here?
        # Fetching here is safer but slower. 
        # GAS onEdit gives us the range. 
        # Ideally, GAS sends the header name to save a roundtrip.

        source_header = event_data.get("header_name")
        if not source_header:
            logger.warning("sync_event_no_header", sheet=sheet_name)
            # Try to fetch header from row 1 using col index
            try:
                ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
                # fetch specific cell or row? row 1 is better to cache?
                # For now simple:
                source_header = ws.cell(1, col_idx).value
                logger.info("fetched_missing_header", header=source_header)
            except Exception as e:
                logger.error("fetch_header_failed", error=str(e))
                return {"status": "failed", "reason": "Could not determine header"}

        # 3. Find Matching Rules (с поддержкой bidirectional)
        matching_rules = self._get_matching_rules(rules, sheet_name, source_header)
        
        if not matching_rules:
            # logger.info("no_matching_sync_rules", sheet=sheet_name, header=source_header)
            self._log_to_session(
                spreadsheet_id, 
                "СИНХРОНИЗАЦИЯ", 
                "Проверка правил", 
                f"Правила для колонки '{source_header}' листа '{sheet_name}' не найдены. Синхронизация пропущена.",
                "ℹ️ SKIP"
            )
            return {"status": "skipped", "reason": "No matching rules"}

        self._log_to_session(
            spreadsheet_id, 
            "СИНХРОНИЗАЦИЯ", 
            "Применение правил", 
            f"Найдено {len(matching_rules)} правил для колонки '{source_header}'",
            "🔄 ПРОЦЕСС"
        )

        # 4. Get Row Key (ID)
        row_key = event_data.get("row_key") # Value of Col A
        if not row_key and row_idx and row_idx > 1:
             # Fetch ID if missing
             try:
                 ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
                 row_key = ws.cell(row, 1).value
                 logger.info("fetched_missing_key", key=row_key)
             except Exception:
                 pass
                 
        if not row_key:
             return {"status": "skipped", "reason": "No row key provided"}

        results = []
        for rule in matching_rules:
            try:
                # Определяем целевой лист/заголовок (с учетом bidirectional)
                target_sheet, target_header = self._get_target_for_rule(rule, sheet_name)

                # Блокируем целевую ячейку для защиты от циклов
                self._lock_cell(spreadsheet_id, row_key, target_header)

                res = self._apply_rule_extended(
                    spreadsheet_id,
                    rule,
                    row_key,
                    new_value,
                    sheet_name,
                    project=event_data.get("project"),
                )
                results.append(res)

                # Log success for each rule
                if rule.mode == "bidirectional":
                    log_msg = f"{rule.sheet_a}#{rule.header_a} <-> {rule.sheet_b}#{rule.header_b} (ID={row_key})"
                else:
                    log_msg = f"{rule.source_sheet}#{rule.source_header} -> {rule.target_sheet}#{rule.target_header} (ID={row_key})"

                self._log_to_session(
                    spreadsheet_id,
                    "СИНХРОНИЗАЦИЯ",
                    "Синхронизация выполнена",
                    log_msg,
                    "✅ OK"
                )
            except Exception as e:
                logger.error("rule_application_failed", error=str(e), rule=rule.id)
                results.append({"rule_id": rule.id, "status": "failed", "error": str(e)})
                self._log_to_session(
                    spreadsheet_id,
                    "СИНХРОНИЗАЦИЯ",
                    "Ошибка правила",
                    f"Не удалось применить правило {rule.id}: {str(e)}",
                    "❌ ОШИБКА"
                )

        # 5. Handle Cascades (Certification, etc)
        # Verify arguments to avoid TypeError
        if source_header:
            try:
                self._handle_cascades(spreadsheet_id, sheet_name, row, source_header)
            except Exception as e:
                logger.error("cascade_invocation_failed", error=str(e))
                # Do not re-raise to avoid blocking response

        # 6. Handle Deadline Autofill (Migration from GAS)
        if source_header:
            try:
                self._check_and_update_deadlines(spreadsheet_id, sheet_name, row, source_header)
            except Exception as e:
                logger.error("deadline_autofill_failed", error=str(e))

        result = {"status": "processed", "rules_matched": len(matching_rules), "results": results}

        # Log summary to Sheets if available
        if self.logging_service and len(results) > 0:
            success_count = sum(1 for r in results if r.get("status") == "success")
            self.logging_service.log_sync_summary(
                spreadsheet_id=spreadsheet_id,
                source_sheet=sheet_name,
                target_sheet=",".join([r.get("target_sheet", "unknown") for r in results if "target_sheet" in r][:3]) or "multiple",
                status="success" if success_count > 0 else "error",
                details=f"Обработано {len(results)} правил, успешно {success_count}",
            )

        return result

    def _log_sync_journal(
        self,
        spreadsheet_id: str,
        row_key: str,
        source_info: str,
        target_info: str,
        old_val: str,
        new_val: str,
        category: str,
        status: str,
        project: Optional[str] = None,
        rule_id: Optional[str] = None,
    ):
        if not self.sync_log_service:
            return

        try:
            self.sync_log_service.add_entry(
                spreadsheet_id=spreadsheet_id,
                project=project,
                row_key=row_key,
                source_info=source_info,
                target_info=target_info,
                old_value=str(old_val),
                new_value=str(new_val),
                category=category,
                status=status,
                rule_id=rule_id,
            )
        except Exception as e:
            logger.error("sync_journal_write_failed", error=str(e))

    def _apply_rule(
        self,
        spreadsheet_id: str,
        rule: Any,
        row_key: str,
        new_value: Any,
        project: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Apply a single sync rule."""
        target_ss_id = rule.target_doc_id if rule.is_external else spreadsheet_id
        
        # 1. Connect to target sheet
        try:
            ws = self.sheets_service.get_worksheet(target_ss_id, rule.target_sheet)
        except Exception:
            return {"rule_id": rule.id, "status": "failed", "error": "Target sheet not found"}

        # 2. Find target column index by header
        headers = self.sheets_service.get_worksheet_headers(target_ss_id, rule.target_sheet)
        try:
            target_col = headers.index(rule.target_header) + 1
        except ValueError:
             return {"rule_id": rule.id, "status": "failed", "error": f"Target header '{rule.target_header}' not found"}

        # 3. Find target row by Key (Col A)
        target_row = self.sheets_service.get_row_by_id(target_ss_id, rule.target_sheet, row_key)
        if not target_row:
             return {"rule_id": rule.id, "status": "skipped", "reason": "Key not found in target"}
        
        # 4. Update Value
        # Optimization: We skip fetching old_val to save one API call. 
        # If logging specifically needs it, we could fetch, but for speed we just update.
        old_val = "[RESTORED_FROM_LOGS]" # Placeholder or we can fetch only if log_level is DEBUG
        
        ws.update_cell(target_row, target_col, new_value)
        
        # 5. Log to sync journal (server)
        self._log_sync_journal(
            spreadsheet_id, 
            row_key=row_key,
            source_info=f"Rule: {rule.category}", 
            target_info=f"{rule.target_sheet}!{rule.target_header} (Row {target_row})",
            old_val=str(old_val),
            new_val=str(new_value),
            category=rule.category,
            status="SUCCESS",
            project=project,
            rule_id=rule.id if hasattr(rule, "id") else None,
        )
        
        # 6. Log to session log (new)
        self._log_to_session(
            spreadsheet_id,
            "АВТОМАТИКА",
            f"Правило выполнено: {rule.category}",
            f"Передано: {rule.target_sheet}!{rule.target_header} (Row {target_row}) | {old_val} -> {new_value}",
            "✅ УСПЕХ"
        )
        
        return {"rule_id": rule.id, "status": "success"}

    def _apply_rule_extended(
        self,
        spreadsheet_id: str,
        rule: SyncRule,
        row_key: str,
        new_value: Any,
        source_sheet: str,
        project: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Apply a sync rule with bidirectional support."""
        # Определяем целевой лист и заголовок
        target_sheet, target_header = self._get_target_for_rule(rule, source_sheet)

        if not target_sheet or not target_header:
            return {"rule_id": rule.id, "status": "failed", "error": "Invalid target configuration"}

        target_ss_id = rule.target_doc_id if rule.is_external else spreadsheet_id

        # 1. Connect to target sheet
        try:
            ws = self.sheets_service.get_worksheet(target_ss_id, target_sheet)
        except Exception:
            return {"rule_id": rule.id, "status": "failed", "error": f"Target sheet '{target_sheet}' not found"}

        # 2. Find target column index by header
        headers = self.sheets_service.get_worksheet_headers(target_ss_id, target_sheet)
        try:
            target_col = headers.index(target_header) + 1
        except ValueError:
            return {"rule_id": rule.id, "status": "failed", "error": f"Target header '{target_header}' not found"}

        # 3. Find target row by Key (Col A)
        target_row = self.sheets_service.get_row_by_id(target_ss_id, target_sheet, row_key)
        if not target_row:
            return {"rule_id": rule.id, "status": "skipped", "reason": "Key not found in target"}

        # 4. Update Value
        old_val = "[OPTIMIZED]"
        ws.update_cell(target_row, target_col, new_value)

        # 5. Log to sync journal
        self._log_sync_journal(
            spreadsheet_id,
            row_key=row_key,
            source_info=f"Rule: {rule.category} ({rule.mode})",
            target_info=f"{target_sheet}!{target_header} (Row {target_row})",
            old_val=str(old_val),
            new_val=str(new_value),
            category=rule.category,
            status="SUCCESS",
            project=project,
            rule_id=rule.id,
        )

        # 6. Log to session log
        self._log_to_session(
            spreadsheet_id,
            "АВТОМАТИКА",
            f"Правило выполнено: {rule.category}",
            f"Передано: {target_sheet}!{target_header} (Row {target_row}) | {old_val} -> {new_value}",
            "✅ УСПЕХ"
        )

        return {"rule_id": rule.id, "status": "success"}

    def _log_to_session(self, spreadsheet_id: str, category: str, action: str, details: str, status: str):
        """Helper to log to the session sheet if available."""
        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, category, action, details, status)

    def _handle_cascades(self, spreadsheet_id: str, sheet_name: str, row_idx: int, col_name: str):
        """
        Handle specific business logic cascades.
        """
        if sheet_name == "Сертификация":
            self._process_certification_cascade(spreadsheet_id, sheet_name, row_idx, col_name)

    def _process_certification_cascade(self, spreadsheet_id: str, sheet_name: str, row_idx: int, updated_header: str):
        """
        Replicates _runCertificationCascade from GAS.
        Handles: "Наименование ДС", "Объём англ.", "Наименование для инвойса", etc.
        """
        key = updated_header.lower().strip()
        triggers = {
            "наименования рус по дс",
            "наименования англ по дс",
            "объём",
            "код тн вэд"
        }
        
        # Check if updated header is a trigger
        if key not in triggers:
            return

        self._log_to_session(
            spreadsheet_id, 
            "КАСКАД", 
            "Запуск каскадного обновления", 
            f"Триггер: '{updated_header}' на листе '{sheet_name}' Row={row_idx}",
            "🔄"
        )

        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
            headers = ws.row_values(1)
            # Map header name to 1-based index
            header_map = {h.lower().strip(): i + 1 for i, h in enumerate(headers)}
            
            # Helper to get/set
            def get_val(name):
                idx = header_map.get(name.lower().strip())
                if not idx: return None
                val = ws.cell(row_idx, idx).value
                return str(val).strip() if val else ""
            
            def set_val(name, val):
                idx = header_map.get(name.lower().strip())
                if idx:
                    current = ws.cell(row_idx, idx).value
                    if str(current or "").strip() != str(val).strip():
                        ws.update_cell(row_idx, idx, val)
                        logger.info("cascade_update", sheet=sheet_name, row=row_idx, field=name, val=val)
                        if hasattr(self, 'logging_service') and self.logging_service:
                             self.logging_service.add_log(spreadsheet_id, "КАСКАД", "Поле обновлено", f"{name}: '{current}' -> '{val}'", "✅")

            # 1. Fetch Source Values
            rus_name = get_val("Наименования рус по ДС") or ""
            eng_name = get_val("Наименования англ по ДС") or ""
            volume = get_val("Объём") or ""
            tnved = get_val("Код ТН ВЭД") or ""
            current_vol_en = get_val("Объём англ.") or ""

            # 2. Compute "Наименование ДС"
            # Logic: If rus_name ok and eng_name ok. If rus_name ends with comma, join with space, else " / "
            new_ds_name = ""
            if rus_name and eng_name:
                sep = " " if rus_name.strip().endswith(",") else " / "
                new_ds_name = f"{rus_name}{sep}{eng_name}"
            else:
                new_ds_name = rus_name or eng_name

            # 3. Compute "Объём англ." (only if volume was the trigger)
            new_vol_en = current_vol_en
            if key == "объём":
                new_vol_en = volume
                replacements = [
                    ("мл", "ml"),
                    ("гр", "g"),
                    ("Тестер", "Tester"),
                    ("шт. х", "*")
                ]
                for old, new in replacements:
                    # case-insensitive replace
                    import re
                    new_vol_en = re.sub(re.escape(old), new, new_vol_en, flags=re.IGNORECASE)
                new_vol_en = " ".join(new_vol_en.split()) # normalize spaces

            # 4. Compute "Наименование для инвойса" (INV_RU)
            # Logic: DS_NAME + " " + VOL + (if tnved: \nКод ТН ВЭД: ...)
            new_inv_ru = f"{new_ds_name} {volume}".strip()
            if tnved:
                new_inv_ru += f"\nКод ТН ВЭД: {tnved}"

            # 5. Compute "Наименование для инвойса Англ" (INV_EN)
            # Logic: EngName + " " + VolEn + (if tnved: \nCode: ...)
            # Determine which VolEn to use? Logic says "new_vol_en || current_vol_en"
            vol_en_to_use = new_vol_en if new_vol_en else current_vol_en
            new_inv_en = f"{eng_name} {vol_en_to_use}".strip()
            if tnved:
                new_inv_en += f"\nCode: {tnved}"

            # 6. Write Changes
            if rus_name or eng_name:
                set_val("Наименование ДС", new_ds_name)
            
            if key == "объём":
                set_val("Объём англ.", new_vol_en)
            
            set_val("Наименование для инвойса", new_inv_ru)
            set_val("Наименование для инвойса Англ", new_inv_en)

        except Exception as e:
            logger.error("cascade_failed", error=str(e), sheet=sheet_name)

    async def add_article(self, spreadsheet_id: str, article: Optional[str] = None, project: str = "UNKNOWN") -> Dict[str, Any]:
        """
        Add a new article to all relevant sheets. 
        If article is empty, generates next ID based on project prefix.
        """
        # 1. Resolve Project and Prefix
        if project == "UNKNOWN":
            # Simple heuristic: we can look into cached rules or just try to find prefix in sheet A1
            # But usually project is passed from GAS. 
            # If not, we try to detect from Spreadsheet name or just use PROJECT_KEY if defined.
            pass

        prefix = self._get_project_prefix(project)
        
        # 2. Generate ID if not provided
        if not article or str(article).strip() == "":
            if self.logging_service:
                self.logging_service.add_log(spreadsheet_id, "АРТИКУЛ", "Определение нового ID", f"Проект: {project}, Префикс: {prefix}", "🔄")
            article = self._get_next_id(spreadsheet_id, prefix)
            logger.info("auto_generated_article_id", id=article, project=project)
            if self.logging_service:
                self.logging_service.add_log(spreadsheet_id, "АРТИКУЛ", "ID определен", f"Новый ID: {article}", "✅")
        
        article = str(article).strip()
        
        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "АРТИКУЛ", f"Добавление артикула {article}", "Начало процесса создания строк", "🚀")

        TARGET_SHEETS = [
            "Заказ", 
            "Этикетки", 
            "Сертификация", 
            "Динамика цены", 
            "Расчет цены",
            "ABC-Анализ",
            "New sert"
        ]
        
        results = {}
        
        for sheet_name in TARGET_SHEETS:
            try:
                ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
                
                # Check if exists (Col A)
                col_a = ws.col_values(1)
                if article in col_a:
                    results[sheet_name] = "Exists"
                    continue
                    
                # Append row
                # We append [article] and let other cols be empty
                ws.append_row([article])
                results[sheet_name] = "Added"
                if self.logging_service:
                    self.logging_service.add_log(spreadsheet_id, "АРТИКУЛ", "Создание строки", f"Лист: {sheet_name}", "✅ Added")
                
            except Exception as e:
                logger.error("add_article_failed_sheet", sheet=sheet_name, error=str(e))
                results[sheet_name] = f"Error: {str(e)}"
                if self.logging_service:
                    self.logging_service.add_log(spreadsheet_id, "АРТИКУЛ", f"Ошибка на листе {sheet_name}", str(e), "❌ ERR")
                
        if self.logging_service:
             self.logging_service.add_log(spreadsheet_id, "АРТИКУЛ", f"Процесс завершен: {article}", f"Результатов: {len(results)}", "✅ ГОТОВО")

        return {"status": "success", "article": article, "details": results}

    def delete_articles(self, spreadsheet_id: str, articles: List[str]) -> Dict[str, Any]:
        """
        Delete rows with matching articles from all relevant sheets.
        """
        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "УДАЛЕНИЕ", f"Запуск удаления артов: {len(articles)} шт.", f"Список: {', '.join(articles[:5])}...", "🔄")

        # Same list as add_article for now, mirroring the ecosystem
        TARGET_SHEETS = [
            "Заказ", 
            "Этикетки", 
            "Сертификация", 
            "Динамика цены", 
            "Расчет цены",
            "ABC-Анализ",
            "New sert"
        ]
        
        results = {}
        deleted_count = 0
        
        for sheet_name in TARGET_SHEETS:
            try:
                ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
                
                # Fetch Col A (IDs)
                col_a = ws.col_values(1)
                
                # Find rows to delete (bottom-up to preserve indices)
                rows_to_delete = [] # list of (index (1-based), article)
                
                for i, val in enumerate(col_a):
                    if val in articles:
                        rows_to_delete.append(i + 1)
                
                if not rows_to_delete:
                    results[sheet_name] = "No matches"
                    continue
                    
                # Delete bottom-up
                for row_idx in sorted(rows_to_delete, reverse=True):
                    ws.delete_rows(row_idx)
                    
                results[sheet_name] = f"Deleted {len(rows_to_delete)} rows"
                deleted_count += len(rows_to_delete)
                if self.logging_service:
                    self.logging_service.add_log(spreadsheet_id, "УДАЛЕНИЕ", f"Лист {sheet_name}", f"Удалено {len(rows_to_delete)} строк", "✅")
                
            except Exception as e:
                logger.error("delete_articles_failed_sheet", sheet=sheet_name, error=str(e))
                results[sheet_name] = f"Error: {str(e)}"
                if self.logging_service:
                    self.logging_service.add_log(spreadsheet_id, "УДАЛЕНИЕ", f"Ошибка на листе {sheet_name}", str(e), "❌ ERR")
                
        if self.logging_service:
             self.logging_service.add_log(spreadsheet_id, "УДАЛЕНИЕ", "Процесс завершен", f"Всего удалено {deleted_count} строк в {len(results)} листах", "✅ ГОТОВО")

        return {"status": "success", "total_deleted": deleted_count, "details": results}

    def _format_expiry_date(self, value: Any) -> str:
        if not value: return ""
        s = str(value).strip()
        if not s: return ""
        import re
        match = re.match(r"^(\d{1,2})[\./](\d{4})$", s)
        if match:
             return f"{int(match.group(1)):02d}.{match.group(2)}"
        return s

    def _check_and_update_deadlines(self, spreadsheet_id: str, sheet_name: str, row_idx: int, header_changed: str):
        if sheet_name != "Заказ": return
        triggers = {"сг 1", "сг 2", "сг 3"}
        if str(header_changed).lower().strip() not in triggers: return

        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
            headers = [str(h or "").strip() for h in ws.row_values(1)]
            normalized_headers = {h.lower(): i + 1 for i, h in enumerate(headers)} # 1-based cols

            def get_val(name):
                col = normalized_headers.get(name.lower())
                if not col: return ""
                val = ws.cell(row_idx, col).value
                return str(val).strip() if val else ""

            def set_val(name, val):
                col = normalized_headers.get(name.lower())
                if col: ws.update_cell(row_idx, col, val)

            dates = []
            for h in ["СГ 1", "СГ 2", "СГ 3"]:
                d = self._format_expiry_date(get_val(h))
                if d: dates.append(d)
            
            res = "\n".join(dates)
            set_val("Срок#", res)
            set_val("Срок", res)
            
            if self.logging_service:
                 self.logging_service.add_log(spreadsheet_id, "АВТОЗАПОЛНЕНИЕ", "Обновление сроков", f"Строка {row_idx}: {res}", "✅")
        except Exception as e:
            logger.error("deadline_update_error", error=str(e))

    # === CRUD методы для управления правилами ===

    def _load_rules_raw(self, spreadsheet_id: str) -> List[SyncRule]:
        """Load all rules (including disabled) without caching."""
        rules_path = os.path.join("config", "rules", f"{spreadsheet_id}.yaml")
        if not os.path.exists(rules_path):
            return []

        with open(rules_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or []

        rules = []
        for row in data:
            try:
                mode = str(row.get("mode", "unidirectional")).strip() or "unidirectional"
                rule = SyncRule(
                    id=str(row.get("id", "")).strip(),
                    enabled=bool(row.get("enabled", True)),
                    category=str(row.get("category", "")).strip(),
                    hashtags=str(row.get("hashtags", "")).strip(),
                    mode=mode,
                    source_sheet=str(row.get("source_sheet", "")).strip(),
                    source_header=str(row.get("source_header", "")).strip(),
                    target_sheet=str(row.get("target_sheet", "")).strip(),
                    target_header=str(row.get("target_header", "")).strip(),
                    sheet_a=str(row.get("sheet_a", "")).strip() or None,
                    header_a=str(row.get("header_a", "")).strip() or None,
                    sheet_b=str(row.get("sheet_b", "")).strip() or None,
                    header_b=str(row.get("header_b", "")).strip() or None,
                    is_external=bool(row.get("is_external", False)),
                    target_doc_id=(str(row.get("target_doc_id", "")).strip() or None)
                )
                rules.append(rule)
            except Exception:
                continue
        return rules

    def _make_bidirectional_id(self, idx: int, sheet_a: str, sheet_b: str, header: str) -> str:
        """Generate ID for bidirectional rule."""
        code_a = self.sheet_codes.get(sheet_a, (sheet_a[:1] or "?"))
        code_b = self.sheet_codes.get(sheet_b, (sheet_b[:1] or "?"))
        return f"{idx:03d}-{code_a}<->{code_b}({header})"

    def _validate_rule_model(self, rule: SyncRule) -> None:
        mode = (rule.mode or "unidirectional").strip()
        if mode not in {"unidirectional", "bidirectional"}:
            raise ValueError(f"Неизвестный режим правила: {mode}")

        if mode == "bidirectional":
            if rule.is_external:
                raise ValueError("Bidirectional правила не поддерживают внешние документы")
            if not rule.sheet_a or not rule.header_a or not rule.sheet_b or not rule.header_b:
                raise ValueError("Bidirectional правило требует sheet_a/header_a и sheet_b/header_b")
        else:
            if not rule.source_sheet or not rule.source_header or not rule.target_sheet or not rule.target_header:
                raise ValueError("Правило требует source_sheet/source_header/target_sheet/target_header")
            if rule.is_external and not rule.target_doc_id:
                raise ValueError("Для внешнего правила нужен target_doc_id")

    def create_rule(self, spreadsheet_id: str, rule_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a new sync rule."""
        rules = self._load_rules_raw(spreadsheet_id)
        idx = len(rules) + 1
        mode = str(rule_data.get("mode", "unidirectional")).strip() or "unidirectional"
        if mode not in {"unidirectional", "bidirectional"}:
            raise ValueError(f"Неизвестный режим правила: {mode}")

        # Generate ID
        if mode == "bidirectional":
            sheet_a = str(rule_data.get("sheet_a", "")).strip()
            sheet_b = str(rule_data.get("sheet_b", "")).strip()
            header = str(rule_data.get("header_a", "")).strip()
            if not sheet_a or not sheet_b or not header:
                raise ValueError("Bidirectional правило требует sheet_a, sheet_b и header_a")
            if bool(rule_data.get("is_external", False)):
                raise ValueError("Bidirectional правила не поддерживают внешние документы")
            rule_id = self._make_bidirectional_id(idx, sheet_a, sheet_b, header)
        else:
            src = str(rule_data.get("source_sheet", "")).strip()
            tgt = str(rule_data.get("target_sheet", "")).strip()
            header = str(rule_data.get("source_header", "")).strip()
            if not src or not tgt or not header:
                raise ValueError("Правило требует source_sheet, source_header и target_sheet")
            rule_id = self._make_rule_id(idx, src, tgt, header)

        new_rule = SyncRule(
            id=rule_id,
            enabled=bool(rule_data.get("enabled", True)),
            category=str(rule_data.get("category", "")).strip(),
            hashtags=str(rule_data.get("hashtags", "")).strip(),
            mode=mode,
            source_sheet=str(rule_data.get("source_sheet", "")).strip(),
            source_header=str(rule_data.get("source_header", "")).strip(),
            target_sheet=str(rule_data.get("target_sheet", "")).strip(),
            target_header=str(rule_data.get("target_header", "")).strip(),
            sheet_a=str(rule_data.get("sheet_a", "")).strip() or None,
            header_a=str(rule_data.get("header_a", "")).strip() or None,
            sheet_b=str(rule_data.get("sheet_b", "")).strip() or None,
            header_b=str(rule_data.get("header_b", "")).strip() or None,
            is_external=bool(rule_data.get("is_external", False)),
            target_doc_id=str(rule_data.get("target_doc_id", "")).strip() or None
        )
        self._validate_rule_model(new_rule)

        rules.append(new_rule)
        self._save_rules_yaml(spreadsheet_id, rules)

        # Invalidate cache
        self._rules_cache = []
        self._rules_cache_time = 0

        logger.info("rule_created", rule_id=rule_id, spreadsheet_id=spreadsheet_id)
        return new_rule.model_dump()

    def update_rule(self, spreadsheet_id: str, rule_id: str, updates: Dict[str, Any]) -> Dict[str, Any]:
        """Update an existing rule (partial update)."""
        rules = self._load_rules_raw(spreadsheet_id)

        for i, rule in enumerate(rules):
            if rule.id == rule_id:
                rule_dict = rule.model_dump()
                rule_dict.update(updates)
                rules[i] = SyncRule(**rule_dict)
                self._validate_rule_model(rules[i])
                self._save_rules_yaml(spreadsheet_id, rules)

                # Invalidate cache
                self._rules_cache = []
                self._rules_cache_time = 0

                logger.info("rule_updated", rule_id=rule_id, spreadsheet_id=spreadsheet_id)
                return rules[i].model_dump()

        raise ValueError(f"Rule not found: {rule_id}")

    def delete_rule(self, spreadsheet_id: str, rule_id: str) -> bool:
        """Delete a rule by ID."""
        rules = self._load_rules_raw(spreadsheet_id)
        original_count = len(rules)
        rules = [r for r in rules if r.id != rule_id]

        if len(rules) == original_count:
            raise ValueError(f"Rule not found: {rule_id}")

        self._save_rules_yaml(spreadsheet_id, rules)

        # Invalidate cache
        self._rules_cache = []
        self._rules_cache_time = 0

        logger.info("rule_deleted", rule_id=rule_id, spreadsheet_id=spreadsheet_id)
        return True

    def toggle_rule(self, spreadsheet_id: str, rule_id: str, enabled: bool) -> Dict[str, Any]:
        """Toggle rule enabled/disabled status."""
        return self.update_rule(spreadsheet_id, rule_id, {"enabled": enabled})

    async def replicate_new_articles(self, spreadsheet_id: str, new_articles: List[Dict[str, Any]]):
        """
        Replicate new articles across all sheets and apply smart match data.
        """
        if not new_articles:
            return

        logger.info(f"replicating_{len(new_articles)}_articles", spreadsheet_id=spreadsheet_id)

        TARGET_SHEETS = [
            "Заказ", 
            "Этикетки", 
            "Сертификация", 
            "Динамика цены", 
            "Расчет цены",
            "ABC-Анализ",
            "New sert"
        ]

        for sheet_name in TARGET_SHEETS:
            try:
                ws = self.sheets_service.get_worksheet(spreadsheet_id, sheet_name)
                
                # Fetch Col A to avoid duplicates
                col_a = ws.col_values(1)
                existing_ids = {str(val).strip() for val in col_a if val}
                
                rows_to_append = []
                for art_data in new_articles:
                    article_id = art_data["id"]
                    if article_id not in existing_ids:
                        rows_to_append.append([article_id])
                
                if rows_to_append:
                    ws.append_rows(rows_to_append)
                    logger.info("replicated_articles", sheet=sheet_name, count=len(rows_to_append))
                    if self.logging_service:
                        self.logging_service.add_log(spreadsheet_id, "АРТИКУЛ", f"Лист {sheet_name}", f"Добавлено строк: {len(rows_to_append)}", "✅")
                
                # Special handling for Сертификация: Apply Smart Match
                if sheet_name == "Сертификация":
                    # Refresh Col A to get correct row indices
                    updated_col_a = ws.col_values(1)
                    idx_map = {str(val).strip(): i + 1 for i, val in enumerate(updated_col_a) if val}
                    
                    headers = ws.row_values(1)
                    header_map = {h.strip().lower(): i + 1 for i, h in enumerate(headers) if h}
                    
                    # Fields to fill from match_info
                    fields_to_fill = {
                        "категория": "категория",
                        "вид продукции": "Вид продукции",
                        "код тн вэд": "Код ТН ВЭД",
                        "наименования англ по дс": "Наименования англ по ДС",
                        "наименования рус по дс": "Наименования рус по ДС",
                        "объём": "Объём",
                        "дс": "ДС",
                        "номер дс": "номер ДС",
                        "дс до": "ДС до",
                        "протокол дс": "Протокол ДС",
                        "протокол доп": "Протокол доп",
                        "inci": "INCI",
                        "coa": "COA"
                    }
                    
                    updates = []
                    for art_data in new_articles:
                        article_id = art_data["id"]
                        match_info = art_data.get("match_info")
                        if not match_info or not match_info.get("match_found"):
                            continue
                            
                        details = match_info.get("matched_product_details", {})
                        row_idx = idx_map.get(article_id)
                        if not row_idx:
                            continue
                            
                        for header_low, detail_key in fields_to_fill.items():
                            col_idx = header_map.get(header_low)
                            val = details.get(detail_key)
                            if col_idx and val:
                                updates.append({
                                    'range': rowcol_to_a1(row_idx, col_idx),
                                    'values': [[val]]
                                })
                        
                    if updates:
                        # Split updates into chunks to avoid too large requests
                        chunk_size = 50
                        for i in range(0, len(updates), chunk_size):
                            ws.batch_update(updates[i:i+chunk_size], value_input_option='USER_ENTERED')
                        
                        logger.info("smart_match_data_applied", count=len(new_articles))
                        if self.logging_service:
                             self.logging_service.add_log(spreadsheet_id, "Smart Match", "Данные сертификации заполнены", f"Записей: {len(new_articles)}", "✅")
                        
                        # Trigger cascades for each updated row
                        for art_data in new_articles:
                            if art_data.get("match_info"):
                                r_idx = idx_map.get(art_data["id"])
                                if r_idx:
                                    self._handle_cascades(spreadsheet_id, sheet_name, r_idx, "Наименования рус по ДС")

            except Exception as e:
                logger.error("replication_failed", sheet=sheet_name, error=str(e), exc_info=True)
                if self.logging_service:
                     self.logging_service.add_log(spreadsheet_id, "АРТИКУЛ", f"Ошибка репликации: {sheet_name}", str(e), "❌")
