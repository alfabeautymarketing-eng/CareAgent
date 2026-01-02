"""
Sorting Service - High-performance sheet sorting with batch operations.
Replaces the slow GAS structureMultipleSheets function.
"""

from typing import Dict, List, Any, Optional, Tuple
from src.services.sheets import SheetsService
from src.utils.logger import logger


class SortingService:
    """
    High-performance sorting service using batch operations.
    Sorts multiple sheets by grouping rows based on ID-G (manufacturer) or ID-L (price line).
    """

    # Sheet configuration
    SHEETS_TO_SORT = ["Заказ", "Динамика цены", "Расчет цены"]
    PRIMARY_SHEET = "Главная"
    PRICE_SHEET = "Прайс"

    # Column mappings
    FIELDS = {
        "ID": "ID",
        "IDP": "ID-P",
        "IDG": "ID-G",
        "IDL": "ID-L",
        "LINE": "Линия",
        "GROUP": "Группа",
        "TITLE": "Название  ENG прайс произв",
        "DS_NAME": "Наименования рус по ДС",
        "GROUP_LINE": "Группа линии",
        "LINE_PRICE": "Линия Прайс",
    }

    # Styling
    GROUP_HEADER_BG_MANUFACTURER = "#666666"
    GROUP_HEADER_BG_PRICE = "#7f6000"
    GROUP_HEADER_FONT = "#FFFFFF"

    def __init__(self):
        self.sheets_service = SheetsService()

    def sort_sheets(self, spreadsheet_id: str, mode: str) -> Dict[str, Any]:
        """
        Sort multiple sheets by manufacturer or price.
        
        Args:
            spreadsheet_id: Google Sheets document ID
            mode: 'byManufacturer' or 'byPrice'
            
        Returns:
            Dict with status and details
        """
        logger.info("sort_sheets_start", spreadsheet_id=spreadsheet_id, mode=mode)
        
        results = {
            "mode": mode,
            "sheets_processed": [],
            "errors": [],
            "total_time_ms": 0,
        }
        
        import time
        start_time = time.time()
        
        try:
            # Step 1: Load all required data in batch
            logger.info("loading_sheet_data")
            sheet_data = self._load_all_sheet_data(spreadsheet_id)
            
            # Step 2: Build group maps from primary/price sheets
            logger.info("building_group_maps")
            group_maps = self._build_group_maps(sheet_data, mode)
            
            # Step 3: Sort each target sheet
            for sheet_name in self.SHEETS_TO_SORT:
                try:
                    logger.info("sorting_sheet", sheet=sheet_name)
                    
                    if sheet_name not in sheet_data:
                        logger.warning("sheet_not_found", sheet=sheet_name)
                        results["errors"].append(f"Sheet '{sheet_name}' not found")
                        continue
                    
                    sorted_data = self._group_and_sort(
                        sheet_data[sheet_name],
                        group_maps,
                        mode
                    )
                    
                    # Write back to sheet
                    self._write_sorted_data(
                        spreadsheet_id,
                        sheet_name,
                        sorted_data,
                        mode
                    )
                    
                    results["sheets_processed"].append({
                        "name": sheet_name,
                        "groups": len(sorted_data["groups"]),
                        "rows": sorted_data["total_rows"],
                    })
                    
                    logger.info("sheet_sorted", sheet=sheet_name, groups=len(sorted_data["groups"]))
                    
                except Exception as e:
                    logger.error("sheet_sort_failed", sheet=sheet_name, error=str(e))
                    results["errors"].append(f"{sheet_name}: {str(e)}")
            
            results["total_time_ms"] = int((time.time() - start_time) * 1000)
            logger.info("sort_sheets_complete", 
                       sheets=len(results["sheets_processed"]),
                       errors=len(results["errors"]),
                       time_ms=results["total_time_ms"])
            
            return results
            
        except Exception as e:
            logger.error("sort_sheets_failed", error=str(e))
            raise

    def _load_all_sheet_data(self, spreadsheet_id: str) -> Dict[str, Dict]:
        """Load all required sheets in batch."""
        gc = self.sheets_service.gc
        spreadsheet = gc.open_by_key(spreadsheet_id)
        
        data = {}
        sheets_to_load = self.SHEETS_TO_SORT + [self.PRIMARY_SHEET, self.PRICE_SHEET]
        
        for sheet_name in sheets_to_load:
            try:
                ws = spreadsheet.worksheet(sheet_name)
                all_values = ws.get_all_values()
                
                if all_values and len(all_values) > 0:
                    headers = [str(h).strip() for h in all_values[0]]
                    header_map = {name: idx for idx, name in enumerate(headers) if name}
                    
                    data[sheet_name] = {
                        "headers": headers,
                        "header_map": header_map,
                        "rows": all_values[1:],  # Skip header row
                        "worksheet": ws,
                    }
                    logger.debug("sheet_loaded", sheet=sheet_name, rows=len(all_values)-1)
            except Exception as e:
                logger.warning("sheet_load_failed", sheet=sheet_name, error=str(e))
        
        return data

    def _build_group_maps(self, sheet_data: Dict, mode: str) -> Dict[str, Any]:
        """Build ID → group info mappings from primary and price sheets."""
        maps = {
            "id_to_group": {},     # ID → {idg, idl, group, line}
            "id_to_price": {},     # ID → {idl, group_line, line_price}
            "idl_to_line": {},     # ID-L → {group_line, line_price}
        }
        
        # Build map from Primary sheet
        if self.PRIMARY_SHEET in sheet_data:
            primary = sheet_data[self.PRIMARY_SHEET]
            hm = primary["header_map"]
            
            id_idx = hm.get(self.FIELDS["ID"])
            idg_idx = hm.get(self.FIELDS["IDG"])
            idl_idx = hm.get(self.FIELDS["IDL"])
            group_idx = hm.get(self.FIELDS["GROUP"])
            line_idx = hm.get(self.FIELDS["LINE"])
            
            if id_idx is not None:
                for row in primary["rows"]:
                    if self._is_row_empty(row):
                        continue
                    
                    id_val = str(row[id_idx]).strip() if id_idx < len(row) else ""
                    if not id_val:
                        continue
                    
                    maps["id_to_group"][id_val] = {
                        "idg": row[idg_idx] if idg_idx is not None and idg_idx < len(row) else None,
                        "idl": row[idl_idx] if idl_idx is not None and idl_idx < len(row) else None,
                        "group": row[group_idx] if group_idx is not None and group_idx < len(row) else None,
                        "line": row[line_idx] if line_idx is not None and line_idx < len(row) else None,
                    }
        
        # Build map from Price sheet (for byPrice mode)
        if mode == "byPrice" and self.PRICE_SHEET in sheet_data:
            price = sheet_data[self.PRICE_SHEET]
            hm = price["header_map"]
            
            id_idx = hm.get(self.FIELDS["ID"])
            idl_idx = hm.get(self.FIELDS["IDL"])
            group_line_idx = hm.get(self.FIELDS["GROUP_LINE"])
            line_price_idx = hm.get(self.FIELDS["LINE_PRICE"])
            
            if id_idx is not None and idl_idx is not None:
                for row in price["rows"]:
                    if self._is_row_empty(row):
                        continue
                    
                    id_val = str(row[id_idx]).strip() if id_idx < len(row) else ""
                    if not id_val:
                        continue
                    
                    idl_raw = row[idl_idx] if idl_idx < len(row) else None
                    group_line = str(row[group_line_idx]).strip() if group_line_idx is not None and group_line_idx < len(row) else ""
                    line_price = str(row[line_price_idx]).strip() if line_price_idx is not None and line_price_idx < len(row) else ""
                    
                    maps["id_to_price"][id_val] = {
                        "idl": idl_raw,
                        "group_line": group_line,
                        "line_price": line_price,
                    }
                    
                    # Build ID-L → line info map
                    if idl_raw and self._is_valid_number(idl_raw):
                        idl_num = int(float(idl_raw))
                        if idl_num not in maps["idl_to_line"]:
                            maps["idl_to_line"][idl_num] = {
                                "group_line": group_line,
                                "line_price": line_price,
                            }
        
        logger.info("group_maps_built",
                   id_to_group=len(maps["id_to_group"]),
                   id_to_price=len(maps["id_to_price"]),
                   idl_to_line=len(maps["idl_to_line"]))
        
        return maps

    def _group_and_sort(self, sheet_data: Dict, group_maps: Dict, mode: str) -> Dict:
        """Group rows and sort within groups."""
        headers = sheet_data["headers"]
        header_map = sheet_data["header_map"]
        rows = sheet_data["rows"]
        
        id_idx = header_map.get(self.FIELDS["ID"])
        idp_idx = header_map.get(self.FIELDS["IDP"])
        ds_name_idx = header_map.get(self.FIELDS["DS_NAME"])
        title_idx = header_map.get(self.FIELDS["TITLE"], 0)
        
        if id_idx is None:
            raise ValueError(f"Column '{self.FIELDS['ID']}' not found")
        
        groups = {}
        
        for row_idx, row in enumerate(rows):
            if self._is_row_empty(row):
                continue
            
            id_val = str(row[id_idx]).strip() if id_idx < len(row) else ""
            if not id_val:
                continue
            
            # Determine group key and title based on mode
            group_key, group_title, sort_value = self._get_group_info(
                id_val, group_maps, mode
            )
            
            # Initialize group if needed
            if group_key not in groups:
                groups[group_key] = {
                    "title": group_title,
                    "sort_value": sort_value,
                    "rows": [],
                }
            
            # Get ID-P for sorting within group
            idp_raw = row[idp_idx] if idp_idx is not None and idp_idx < len(row) else None
            has_idp = self._is_valid_number(idp_raw)
            idp_num = float(idp_raw) if has_idp else None
            
            # Get DS name for fallback sorting
            ds_name = str(row[ds_name_idx]).strip().lower() if ds_name_idx is not None and ds_name_idx < len(row) else ""
            
            groups[group_key]["rows"].append({
                "values": row,
                "has_idp": has_idp,
                "idp": idp_num,
                "ds_name": ds_name,
                "original_idx": row_idx,
            })
        
        # Sort groups
        sorted_group_keys = sorted(groups.keys(), key=lambda k: self._group_sort_key(k, groups[k]))
        
        # Sort rows within each group
        for key in sorted_group_keys:
            group = groups[key]
            if mode == "byPrice":
                # Sort by ID-P ascending, then by original index
                group["rows"].sort(key=lambda r: (0 if r["has_idp"] else 1, r["idp"] or 999999, r["original_idx"]))
            else:
                # byManufacturer: ID-P first, then try to match by DS name
                with_idp = [r for r in group["rows"] if r["has_idp"]]
                without_idp = [r for r in group["rows"] if not r["has_idp"]]
                
                with_idp.sort(key=lambda r: r["idp"])
                without_idp.sort(key=lambda r: r["original_idx"])
                
                # Try to insert without_idp after matching ds_name
                arranged = with_idp.copy()
                for item in without_idp:
                    inserted = False
                    if item["ds_name"]:
                        for i in range(len(arranged) - 1, -1, -1):
                            if arranged[i]["ds_name"] == item["ds_name"]:
                                arranged.insert(i + 1, item)
                                inserted = True
                                break
                    if not inserted:
                        arranged.append(item)
                
                group["rows"] = arranged
        
        # Calculate total rows
        total_rows = sum(len(g["rows"]) for g in groups.values())
        
        return {
            "headers": headers,
            "header_map": header_map,
            "groups": {k: groups[k] for k in sorted_group_keys},
            "title_idx": title_idx,
            "total_rows": total_rows,
        }

    def _get_group_info(self, id_val: str, group_maps: Dict, mode: str) -> Tuple[str, str, Optional[float]]:
        """Get group key, title, and sort value for a row."""
        if mode == "byPrice":
            price_info = group_maps["id_to_price"].get(id_val)
            if price_info:
                idl_raw = price_info.get("idl")
                if self._is_valid_number(idl_raw):
                    idl_num = int(float(idl_raw))
                    line_info = group_maps["idl_to_line"].get(idl_num, {})
                    
                    parts = []
                    if line_info.get("group_line"):
                        parts.append(line_info["group_line"])
                    if line_info.get("line_price"):
                        parts.append(line_info["line_price"])
                    
                    title = "\n".join(parts) if parts else "Группа не определена"
                    return f"LINE_{idl_num}", title, float(idl_num)
            
            return "UNASSIGNED", "Группа не идентифицирована", None
        
        else:  # byManufacturer
            group_info = group_maps["id_to_group"].get(id_val)
            if group_info:
                idg_raw = group_info.get("idg")
                group_raw = group_info.get("group")
                
                if self._is_valid_number(idg_raw):
                    idg_num = int(float(idg_raw))
                    title = str(group_raw).strip() if group_raw else "Группа не определена"
                    return f"GROUP_{idg_num}", title, float(idg_num)
                
                if group_raw:
                    title = str(group_raw).strip()
                    return f"META_{title}", title, None
            
            return "UNASSIGNED", "Группа не идентифицирована", None

    def _group_sort_key(self, key: str, group: Dict) -> Tuple:
        """Generate sort key for groups."""
        if key == "UNASSIGNED":
            return (2, 0, key)  # Last
        if group["sort_value"] is not None:
            return (0, group["sort_value"], key)  # Numeric first
        return (1, 0, key)  # String keys in middle

    def _write_sorted_data(self, spreadsheet_id: str, sheet_name: str, sorted_data: Dict, mode: str):
        """Write sorted data back to sheet with batch operations."""
        gc = self.sheets_service.gc
        spreadsheet = gc.open_by_key(spreadsheet_id)
        ws = spreadsheet.worksheet(sheet_name)
        
        headers = sorted_data["headers"]
        groups = sorted_data["groups"]
        title_idx = sorted_data["title_idx"]
        num_cols = len(headers)
        
        # Clear existing data (keep headers)
        last_row = ws.row_count
        if last_row > 1:
            # Clear from row 2 onwards
            clear_range = f"A2:{self._col_letter(num_cols)}{last_row}"
            ws.batch_clear([clear_range])
        
        # Prepare all data for batch write
        all_rows = []
        header_row_indices = []  # Track which rows are group headers (0-indexed in all_rows)
        
        for group_key, group in groups.items():
            # Add group header row
            header_row = [""] * num_cols
            header_row[title_idx] = group["title"]
            all_rows.append(header_row)
            header_row_indices.append(len(all_rows) - 1)
            
            # Add data rows
            for row_entry in group["rows"]:
                # Ensure row has correct number of columns
                row_values = list(row_entry["values"])
                while len(row_values) < num_cols:
                    row_values.append("")
                all_rows.append(row_values[:num_cols])
        
        # Batch write all data starting from row 2
        if all_rows:
            write_range = f"A2:{self._col_letter(num_cols)}{len(all_rows) + 1}"
            ws.update(write_range, all_rows, value_input_option="USER_ENTERED")
            
            # Apply formatting to header rows (batch)
            bg_color = self.GROUP_HEADER_BG_PRICE if mode == "byPrice" else self.GROUP_HEADER_BG_MANUFACTURER
            
            # Build format requests
            format_requests = []
            for idx in header_row_indices:
                actual_row = idx + 2  # +2 because row 1 is headers, and we're 0-indexed
                format_requests.append({
                    "range": f"A{actual_row}:{self._col_letter(num_cols)}{actual_row}",
                    "format": {
                        "backgroundColor": self._parse_hex_color(bg_color),
                        "textFormat": {
                            "bold": True,
                            "foregroundColor": self._parse_hex_color(self.GROUP_HEADER_FONT),
                        }
                    }
                })
            
            # Apply formatting in batch
            if format_requests:
                ws.batch_format(format_requests)
        
        logger.info("data_written", sheet=sheet_name, rows=len(all_rows), groups=len(groups))

    def _col_letter(self, col_num: int) -> str:
        """Convert column number (1-indexed) to letter."""
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result if result else "A"

    def _parse_hex_color(self, hex_color: str) -> Dict:
        """Convert hex color to RGB dict for gspread."""
        hex_color = hex_color.lstrip("#")
        return {
            "red": int(hex_color[0:2], 16) / 255,
            "green": int(hex_color[2:4], 16) / 255,
            "blue": int(hex_color[4:6], 16) / 255,
        }

    def _is_row_empty(self, row: List) -> bool:
        """Check if a row is empty."""
        return all(cell == "" or cell is None for cell in row)

    def _is_valid_number(self, value) -> bool:
        """Check if value is a valid number."""
        if value is None or value == "":
            return False
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
