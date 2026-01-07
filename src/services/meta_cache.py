import json
import os
import asyncio
from typing import List, Dict, Any, Optional, Set
from datetime import datetime
from src.services.sheets import SheetsService
from src.utils.logger import logger

class MetaCacheService:
    def __init__(self, data_file: str = "data/meta_index.json"):
        self.data_file = data_file
        self.sheets_service = SheetsService()
        self.index = self._load()

    def _load(self) -> Dict[str, Any]:
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logger.error("meta_cache_load_failed", error=str(e))
        return {"updated_at": None, "spreadsheets": {}, "headers": {}}

    def save(self):
        try:
            os.makedirs(os.path.dirname(self.data_file), exist_ok=True)
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.index, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error("meta_cache_save_failed", error=str(e))

    async def index_spreadsheet(self, spreadsheet_id: str):
        """Index all sheets and headers in a spreadsheet."""
        try:
            # We need to get all sheet names first. gspread 'worksheets()' is good.
            sh = self.sheets_service.gc.open_by_key(spreadsheet_id)
            worksheets = sh.worksheets()
            
            ss_entry = {"sheets": {}, "title": sh.title}
            for ws in worksheets:
                sheet_name = ws.title
                try:
                    # Use sheets_service to get headers (it has internal caching too)
                    headers = self.sheets_service.get_worksheet_headers(spreadsheet_id, sheet_name)
                    ss_entry["sheets"][sheet_name] = headers
                    
                    # Update inverted index: header -> list of {ss_id, sheet_name}
                    for h in headers:
                        if not h or not str(h).strip(): continue
                        h_str = str(h).strip()
                        if h_str not in self.index["headers"]:
                            self.index["headers"][h_str] = []
                        
                        # Avoid duplicates
                        exists = any(m["ss"] == spreadsheet_id and m["sheet"] == sheet_name 
                                    for m in self.index["headers"][h_str])
                        if not exists:
                            self.index["headers"][h_str].append({
                                "ss": spreadsheet_id, 
                                "sheet": sheet_name,
                                "ss_title": sh.title
                            })
                except Exception as e:
                    logger.error("index_sheet_failed", ss=spreadsheet_id, sheet=sheet_name, error=str(e))
            
            # Rate limiting: avoid 429 errors
            await asyncio.sleep(2)  # 2 seconds between docs
            
            self.index["spreadsheets"][spreadsheet_id] = ss_entry
            self.index["updated_at"] = datetime.now().isoformat()
            self.save()
            return True
        except Exception as e:
            logger.error("index_ss_failed", ss=spreadsheet_id, error=str(e))
            return False

    async def build_global_index(self, spreadsheet_ids: List[str]):
        """Rebuild index for all provided spreadsheets sequentially."""
        self.index["headers"] = {} # Clear inverted index to avoid stale entries
        
        results = []
        for sid in spreadsheet_ids:
            logger.info("indexing_started", ss=sid)
            res = await self.index_spreadsheet(sid)
            results.append(res)
            # Small delay is already in index_spreadsheet
        
        logger.info("global_index_rebuilt", total=len(spreadsheet_ids), success=sum(results))
        return {
            "total": len(spreadsheet_ids),
            "success": sum(results),
            "updated_at": self.index["updated_at"]
        }

    def find_common_headers(self, ss1: str, s1: str, ss2: str, s2: str) -> List[str]:
        """Fast lookup of common headers between two sheets using the index."""
        h1 = self.index["spreadsheets"].get(ss1, {}).get("sheets", {}).get(s1, [])
        h2 = self.index["spreadsheets"].get(ss2, {}).get("sheets", {}).get(s2, [])
        return [h for h in h1 if h in h2]

    def get_discovery_map(self) -> Dict[str, Any]:
        """Return headers that exist in multiple places across the ecosystem."""
        discovery = {}
        for h, locations in self.index["headers"].items():
            if len(locations) > 1:
                discovery[h] = locations
        return discovery

    def get_indexed_spreadsheets(self) -> List[Dict[str, str]]:
        return [{"id": sid, "title": data.get("title", "Unknown")} 
                for sid, data in self.index["spreadsheets"].items()]
