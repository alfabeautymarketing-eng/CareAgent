import time
import os
import yaml
from pathlib import Path
from typing import List, Dict, Optional, Any, Tuple
from pydantic import BaseModel
from src.services.sheets import SheetsService
from src.utils.logger import logger

class SyncRule(BaseModel):
    id: str
    enabled: bool
    category: str
    hashtags: str = ""
    source_sheet: str
    source_header: str
    target_sheet: str
    target_header: str
    is_external: bool
    target_doc_id: Optional[str] = None

class SyncService:
    def __init__(self, logging_service: Optional[Any] = None):
        self.sheets_service = SheetsService()
        self.logging_service = logging_service
        self._rules_cache: List[SyncRule] = []
        self._rules_cache_time = 0
        self._rules_cache_ttl = 300  # 5 minutes
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

    def _make_rule_id(self, idx: int, source_sheet: str, target_sheet: str, header: str) -> str:
        src_code = self.sheet_codes.get(source_sheet, (source_sheet[:1] or "?"))
        tgt_code = self.sheet_codes.get(target_sheet, (target_sheet[:1] or "?"))
        return f"{idx:03d}-{src_code}-{tgt_code}({header})"

    def _save_rules_yaml(self, spreadsheet_id: str, rules: List[SyncRule]) -> None:
        path = Path("config") / "rules" / f"{spreadsheet_id}.yaml"
        path.parent.mkdir(parents=True, exist_ok=True)
        data = []
        for r in rules:
            d = r.model_dump()
            # exclude hashtags if empty to keep файл компактным
            if not d.get("hashtags"):
                d.pop("hashtags", None)
            data.append(d)
        path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")

    def list_rules(self, spreadsheet_id: str, force_reload: bool = False) -> List[Dict[str, Any]]:
        rules = self._load_rules(spreadsheet_id, force_reload=force_reload)
        return [r.model_dump() for r in rules]

    def save_rules(self, spreadsheet_id: str, rules_payload: List[Dict[str, Any]], validate_headers: bool = True) -> List[Dict[str, Any]]:
        headers_cache: Dict[str, List[str]] = {}
        normalized: List[SyncRule] = []

        for idx, r in enumerate(rules_payload, start=1):
            src_sheet = str(r.get("source_sheet", "")).strip()
            tgt_sheet = str(r.get("target_sheet", "")).strip()
            src_header = str(r.get("source_header", "")).strip()
            tgt_header = str(r.get("target_header", "")).strip()
            category = str(r.get("category", "")).strip()
            enabled = bool(r.get("enabled", True))
            is_external = bool(r.get("is_external", False))
            target_doc_id = str(r.get("target_doc_id", "")).strip() or None

            if not src_sheet or not src_header or not tgt_sheet or not tgt_header:
                raise ValueError(f"Неполное правило (#{idx}): нужны source_sheet, source_header, target_sheet, target_header")

            if validate_headers:
                headers = headers_cache.get(tgt_sheet)
                if headers is None:
                    ws = self.sheets_service.get_worksheet(spreadsheet_id, tgt_sheet)
                    headers = [str(h or "").strip() for h in ws.row_values(1)]
                    headers_cache[tgt_sheet] = headers
                if tgt_header not in headers:
                    raise ValueError(f"Столбец '{tgt_header}' не найден на листе '{tgt_sheet}'")

            rule_id = self._make_rule_id(idx, src_sheet, tgt_sheet, src_header)
            normalized.append(
                SyncRule(
                    id=rule_id,
                    enabled=enabled,
                    category=category,
                    hashtags=str(r.get("hashtags", "")).strip(),
                    source_sheet=src_sheet,
                    source_header=src_header,
                    target_sheet=tgt_sheet,
                    target_header=tgt_header,
                    is_external=is_external,
                    target_doc_id=target_doc_id,
                )
            )

        self._save_rules_yaml(spreadsheet_id, normalized)
        self._rules_cache = normalized
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
        Load sync rules from 'Правила синхро' sheet.
        If config/rules/<spreadsheet_id>.yaml exists, load from server to avoid лишние запросы.
        """
        if not force_reload and (time.time() - self._rules_cache_time < self._rules_cache_ttl) and self._rules_cache:
            return self._rules_cache

        try:
            # 0. Попытка загрузить правила с сервера (YAML) по spreadsheet_id
            rules_path = os.path.join("config", "rules", f"{spreadsheet_id}.yaml")
            if os.path.exists(rules_path):
                with open(rules_path, "r", encoding="utf-8") as f:
                    data = yaml.safe_load(f) or []
                rules = []
                for row in data:
                    try:
                        rule = SyncRule(
                            id=str(row.get("id", "")).strip(),
                            enabled=bool(row.get("enabled", True)),
                            category=str(row.get("category", "")).strip(),
                            hashtags=str(row.get("hashtags", "")).strip(),
                            source_sheet=str(row.get("source_sheet", "")).strip(),
                            source_header=str(row.get("source_header", "")).strip(),
                            target_sheet=str(row.get("target_sheet", "")).strip(),
                            target_header=str(row.get("target_header", "")).strip(),
                            is_external=bool(row.get("is_external", False)),
                            target_doc_id=(str(row.get("target_doc_id", "")).strip() or None)
                        )
                        if not rule.source_sheet or not rule.source_header or not rule.target_sheet or not rule.target_header:
                            continue
                        if rule.is_external and not rule.target_doc_id:
                            continue
                        if not rule.enabled:
                            continue
                        rules.append(rule)
                    except Exception:
                        continue

                if rules:
                    self._rules_cache = rules
                    self._rules_cache_time = time.time()
                    logger.info("sync_rules_loaded_from_server", count=len(rules), source=rules_path)
                    return rules
                # если пустой файл — fallback к листу

            RULES_SHEET = "Правила синхро"
            ws = self.sheets_service.get_worksheet(spreadsheet_id, RULES_SHEET)
            
            # Assuming structure matches GAS: 
            # A=ID, B=Enabled, C=Category, D=Hashtags, E=SourceSheet, F=SourceHeader, 
            # G=TargetSheet, H=TargetHeader, I=External, J=TargetDocId
            
            values = ws.get_all_values()
            if not values or len(values) < 2:
                return []

            rules = []
            # Skip header row (index 0)
            for row in values[1:]:
                # Pad row if too short
                if len(row) < 10:
                    row += [""] * (10 - len(row))
                
                # Check enabled (Column B / index 1)
                enabled_val = row[1].strip().lower()
                enabled = (enabled_val == 'true' or enabled_val == 'да')
                
                if not enabled:
                    continue

                is_external_val = row[8].strip().lower()
                is_external = (is_external_val == 'true' or is_external_val == 'да')

                rule = SyncRule(
                    id=row[0].strip(),
                    enabled=enabled,
                    category=row[2].strip(),
                            hashtags=row[3].strip() if len(row) > 3 else "",
                            source_sheet=row[4].strip(),
                            source_header=row[5].strip(),
                            target_sheet=row[6].strip(),
                            target_header=row[7].strip(),
                            is_external=is_external,
                    target_doc_id=row[9].strip() if row[9].strip() else None
                )
                
                # Validation matches GAS
                if not rule.source_sheet or not rule.source_header or not rule.target_sheet or not rule.target_header:
                    continue
                if rule.is_external and not rule.target_doc_id:
                    continue

                rules.append(rule)

            self._rules_cache = rules
            self._rules_cache_time = time.time()
            logger.info("sync_rules_loaded", count=len(rules))
            return rules

        except Exception as e:
            logger.error("load_rules_failed", error=str(e))
            # Return cached if available even if expired, else empty
            return self._rules_cache

    def sync_row(self, spreadsheet_id: str, sheet_name: str, row_number: int, row_key: str, project: str) -> Dict[str, Any]:
        """
        Sync a single row to target sheets based on rules.
        """
        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "СИНХРО", f"Синхронизация строки {sheet_name}#{row_number}", f"ID={row_key}, Project={project}", "🔄 ЗАПУСК")
            
        rules = self._load_rules(spreadsheet_id)
        effective_rules = [r for r in rules if r.enabled and r.source_sheet == sheet_name]
        
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
            # Get value for this rule's source column
            val = row_data.get(rule.source_header)
            
            if rule.source_header not in row_data:
                logger.warning("source_header_missing", header=rule.source_header)
                if self.logging_service:
                    self.logging_service.add_log(spreadsheet_id, "СИНХРО", "Пропуск правила", f"Исходный заголовок '{rule.source_header}' не найден в строке {row_key} листа {sheet_name}", "⚠️ ПРЕДУПРЕЖДЕНИЕ")
                results.append({"rule_id": rule.id, "status": "skipped", "reason": "Source header missing"})
                continue
                
            res = self._apply_rule(spreadsheet_id, rule, row_key, val)
            results.append(res)
            
        if self.logging_service:
            self.logging_service.add_log(spreadsheet_id, "СИНХРО", "Синхронизация строки завершена", f"Обработано {len(effective_rules)} правил для ID={row_key}. Успешно: {len([r for r in results if r.get('status') == 'success'])}", "✅ УСПЕХ")
        return {"status": "success", "results": results}

    def sync_full(self, spreadsheet_id: str, project: str, source_sheet: str) -> Dict[str, Any]:
        """
        Sync ALL rows in a sheet.
        """
        logger.info("sync_full_start", project=project, sheet=source_sheet)
        
        rules = self._load_rules(spreadsheet_id)
        matching_rules = [r for r in rules if r.source_sheet == source_sheet]
        
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
                    # Get value
                    src_idx = header_map.get(rule.source_header)
                    if src_idx is not None and src_idx < len(row):
                        val = row[src_idx]
                        
                        try:
                            self._apply_rule(spreadsheet_id, rule, article, val)
                        except Exception as e:
                            # Log but don't stop full sync
                            errors.append(f"{article}-{rule.id}: {str(e)}")
                    else:
                        # Header missing or row short
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

    async def sync_event(self, spreadsheet_id: str, event_data: Dict[str, Any]):
        """
        Process a single onEdit event.
        """
        sheet_name = event_data.get("sheet_name")
        row_idx = event_data.get("row")
        col_idx = event_data.get("col")
        source_header = event_data.get("header_name")
        new_value = event_data.get("value")
        row_key = event_data.get("row_key")

        # Пропускаем события на лог-листах
        if sheet_name in {"Логи", "Журнал синхро", "Logs"}:
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

        # 3. Find Matching Rules
        matching_rules = [
            r for r in rules 
            if r.source_sheet == sheet_name and r.source_header == source_header
        ]
        
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
        if not row_key and row > 1:
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
                res = self._apply_rule(spreadsheet_id, rule, row_key, new_value)
                results.append(res)
                
                # Log success for each rule
                self._log_to_session(
                    spreadsheet_id,
                    "СИНХРОНИЗАЦИЯ",
                    "Синхронизация выполнена",
                    f"{rule.source_sheet}#{rule.source_header} -> {rule.target_sheet}#{rule.target_header} (ID={row_key})",
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

        return {"status": "processed", "rules_matched": len(matching_rules), "results": results}

    def _log_to_sheet(self, spreadsheet_id: str, row_key: str, source_info: str, target_info: str, 
                      old_val: str, new_val: str, category: str, hashtags: str, status: str):
        """
        Append log to 'Журнал синхро' with columns:
        Дата/время | ID | Источник | Цель | Старое значение | Новое значение | Категория | Хэштеги | Событие
        """
        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, "Журнал синхро")
            from datetime import datetime
            timestamp = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            ws.append_row([
                timestamp,
                row_key,
                source_info,
                target_info,
                str(old_val),
                str(new_val),
                category,
                hashtags or "",
                status
            ])
        except Exception as e:
            logger.error("log_to_sheet_failed", error=str(e))

    def _apply_rule(self, spreadsheet_id: str, rule: Any, row_key: str, new_value: Any) -> Dict[str, Any]:
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
        
        # 5. Log to sync journal (legacy)
        self._log_to_sheet(
            spreadsheet_id, 
            row_key=row_key,
            source_info=f"Rule: {rule.category}", 
            target_info=f"{rule.target_sheet}!{rule.target_header} (Row {target_row})",
            old_val=str(old_val),
            new_val=str(new_value),
            category=rule.category,
            hashtags=rule.hashtags,
            status="SUCCESS"
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
