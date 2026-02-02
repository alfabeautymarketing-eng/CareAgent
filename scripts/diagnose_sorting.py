import asyncio
import sys
import os
import traceback
from typing import Dict, Any

# Add src to python path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.services.sheets import SheetsService
from src.services.sorting import SortingService
from src.utils.logger import logger

# MOCK PROJECT_IDS if we can't find it
# In a real run we'd import it. 
# For now, I'll ask the user or hardcode a known one if I find it.
# I'll try to discover it from the code imports in the script itself if possible, 
# or just ask the user to provide one as an argument.

def diagnose(spreadsheet_id: str, mode: str = "byPrice"):
    print(f"--- Starting Diagnosis for Spreadsheet: {spreadsheet_id} ---")
    print(f"Mode: {mode}")
    
    sheets_service = SheetsService()
    sorting_service = SortingService()
    
    # 1. Inspect Headers
    print("\n[Step 1] Inspecting Headers...")
    try:
        gc = sheets_service.gc
        sh = gc.open_by_key(spreadsheet_id)
        
        target_sheets = ["Главная", "Прайс"]
        for sheet_name in target_sheets:
            print(f"  Checking sheet '{sheet_name}'...")
            try:
                ws = sh.worksheet(sheet_name)
                headers = ws.row_values(1)
                print(f"    Headers found ({len(headers)}): {headers}")
                
                # Check against expected fields
                missing = []
                # Check logic depends on implementation. 
                # SortingService.FIELDS contains the map.
                # But mapping logic is internal to SortingService.
                # We will check if key columns are present.
                
                # Let's verify commonly required columns for byPrice
                required_cols = []
                if sheet_name == "Главная":
                    required_cols = [
                        SortingService.FIELDS["ID"], 
                        SortingService.FIELDS["IDP"], 
                        SortingService.FIELDS["GROUP"], 
                        SortingService.FIELDS["LINE"]
                    ]
                elif sheet_name == "Прайс":
                    required_cols = [
                        SortingService.FIELDS["ID"], 
                        SortingService.FIELDS["IDL"], 
                        SortingService.FIELDS["GROUP_LINE"], 
                        SortingService.FIELDS["LINE_PRICE"]
                    ]
                
                found_cols = [h.strip() for h in headers]
                for req in required_cols:
                    if req not in found_cols:
                        missing.append(req)
                
                if missing:
                    print(f"    [WARNING] Missing expected columns: {missing}")
                else:
                    print(f"    [OK] All core columns found for logic.")
                    
            except Exception as e:
                print(f"    [ERROR] Could not access sheet '{sheet_name}': {e}")

    except Exception as e:
        print(f"[FATAL] Could not open spreadsheet: {e}")
        return

    # 2. Attempt Sort
    print("\n[Step 2] Attempting Sorting Logic...")
    try:
        # We call the service directly to catch the traceback
        result = sorting_service.sort_sheets(spreadsheet_id, mode)
        print("\n[SUCCESS] Sorting completed without error!")
        print(f"Result: {result}")
    except Exception:
        print("\n[ERROR] Sorting failed with exception:")
        traceback.print_exc()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python diagnose_sorting.py <SPREADSHEET_ID> [MODE]")
        # Try to find a default ID or exit
        # For now, just exit
        sys.exit(1)
    
    sp_id = sys.argv[1]
    mode_arg = sys.argv[2] if len(sys.argv) > 2 else "byPrice"
    
    diagnose(sp_id, mode_arg)
