import time
from typing import List, Dict, Optional, Any
from pydantic import BaseModel
from src.services.sheets import SheetsService
from src.utils.logger import logger

class SyncRule(BaseModel):
    id: str
    enabled: bool
    category: str
    hashtags: str
    source_sheet: str
    source_header: str
    target_sheet: str
    target_header: str
    is_external: bool
    target_doc_id: Optional[str] = None

class SyncService:
    def __init__(self):
        self.sheets_service = SheetsService()
        self._rules_cache: List[SyncRule] = []
        self._rules_cache_time = 0
        self._rules_cache_ttl = 300  # 5 minutes

    def _load_rules(self, spreadsheet_id: str, force_reload: bool = False) -> List[SyncRule]:
        """
        Load sync rules from 'Правила синхро' sheet.
        """
        if not force_reload and (time.time() - self._rules_cache_time < self._rules_cache_ttl) and self._rules_cache:
            return self._rules_cache

        try:
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
                    hashtags=row[3].strip(),
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

    def sync_row(self, spreadsheet_id: str, project: str, article: str, source_sheet: str) -> Dict[str, Any]:
        """
        Sync a single row manually by Article/ID.
        """
        logger.info("sync_row_start", project=project, article=article, sheet=source_sheet)
        
        # 1. Load Rules
        rules = self._load_rules(spreadsheet_id)
        matching_rules = [r for r in rules if r.source_sheet == source_sheet]
        
        if not matching_rules:
            return {"status": "skipped", "reason": "No rules for this sheet"}

        # 2. Get Source Data
        # We need to find the row with this article to get values for all rule columns
        try:
            ws = self.sheets_service.get_worksheet(spreadsheet_id, source_sheet)
            
            # Assume Column 1 is ID/Article
            col_a = ws.col_values(1)
            try:
                row_idx = col_a.index(article) + 1
            except ValueError:
                return {"status": "failed", "error": f"Article '{article}' not found in {source_sheet}"}
                
            # Fetch entire row data (mapped by header)
            # Optimization: could fetch only needed columns, but row is simpler
            row_values = ws.row_values(row_idx)
            headers = ws.row_values(1)
            
            # Map headers to values
            row_data = {}
            for i, h in enumerate(headers):
                if i < len(row_values):
                    row_data[h] = row_values[i]
                    
        except Exception as e:
            logger.error("sync_row_fetch_failed", error=str(e))
            raise

        # 3. Apply Rules
        results = []
        for rule in matching_rules:
            # Get value for this rule's source column
            val = row_data.get(rule.source_header)
            
            # If header exists in source but value might be empty/None, that's fine.
            # If header didn't exist in source map, we skip or error?
            if rule.source_header not in row_data:
                logger.warning("source_header_missing", header=rule.source_header)
                results.append({"rule_id": rule.id, "status": "skipped", "reason": "Source header missing"})
                continue
                
            res = self._apply_rule(spreadsheet_id, rule, article, val)
            results.append(res)
            
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

    def sync_event(self, spreadsheet_id: str, event_data: Dict[str, Any]):
        """
        Process a sync event triggered by an edit.
        event_data: {
            "sheet_name": str,
            "row": int,
            "col": int,
            "value": Any,
            "old_value": Any,
            "user_email": str,
            "ranges": [...] (if needed, but simple cell edit is primary)
        }
        """
        sheet_name = event_data.get("sheet_name")
        row = event_data.get("row")
        col = event_data.get("col")
        new_value = event_data.get("value")
        
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
                source_header = ws.cell(1, col).value
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
            return {"status": "skipped", "reason": "No matching rules"}

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
            except Exception as e:
                logger.error("apply_rule_failed", rule_id=rule.id, error=str(e))
                results.append({"rule_id": rule.id, "status": "failed", "error": str(e)})

        # 5. Handle Cascades (Certification, etc)
        if source_header:
            self._handle_cascades(spreadsheet_id, sheet_name, row, source_header)
        
        return {"status": "processed", "rules_matched": len(matching_rules), "results": results}

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

    def add_article(self, spreadsheet_id: str, article: str) -> Dict[str, Any]:
        """
        Add a new article to all relevant sheets.
        """
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
                
            except Exception as e:
                logger.error("add_article_failed_sheet", sheet=sheet_name, error=str(e))
                results[sheet_name] = f"Error: {str(e)}"
                
        return results

    def delete_articles(self, spreadsheet_id: str, articles: List[str]) -> Dict[str, Any]:
        """
        Delete rows with matching articles from all relevant sheets.
        """
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
                
            except Exception as e:
                logger.error("delete_articles_failed_sheet", sheet=sheet_name, error=str(e))
                results[sheet_name] = f"Error: {str(e)}"
                
        return {"total_deleted": deleted_count, "details": results}

    def _apply_rule(self, current_spreadsheet_id: str, rule: SyncRule, key: str, value: Any):
        """
        Apply a single sync rule.
        """
        target_ss_id = rule.target_doc_id if rule.is_external else current_spreadsheet_id
        
        # 1. Connect to target sheet
        try:
            ws = self.sheets_service.get_worksheet(target_ss_id, rule.target_sheet)
        except Exception:
            return {"rule_id": rule.id, "status": "failed", "error": "Target sheet not found"}

        # 2. Find target column index by header
        # TODO: optimization - cache headers map
        headers = ws.row_values(1)
        try:
            target_col = headers.index(rule.target_header) + 1
        except ValueError:
             return {"rule_id": rule.id, "status": "failed", "error": f"Target header '{rule.target_header}' not found"}

        # 3. Find target row by Key (Col A)
        # TODO: optimization - cache ID map
        col_a = ws.col_values(1)
        try:
            target_row = col_a.index(key) + 1
        except ValueError:
             # Key not found in target, do nothing (according to GAS logic)
             return {"rule_id": rule.id, "status": "skipped", "reason": "Key not found in target"}
        
        # 4. Update Value
        # Check old value to avoid redundant edits?
        # For now, just write. gspread update_cell is simple.
        # Using batch_update or update_cell
        ws.update_cell(target_row, target_col, value)
        
        return {"rule_id": rule.id, "status": "success"}

    def _handle_cascades(self, spreadsheet_id: str, sheet_name: str, row: int, col: int, header: str):
        """
        Handle post-sync logic like Certification cascades.
        """
        # Logic from GAS: if (sheetName === Lib.CONFIG.SHEETS.CERTIFICATION) ...
        # We need strict names.
        
        if sheet_name == "Сертификация":
            # Call certification cascade logic
            pass
        
        # TODO: Add Order sheet overrides if relevant
