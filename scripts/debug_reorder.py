
import os
import sys
from src.services.sheets import SheetsService
from src.api.endpoints import SHEET_ORDER

def debug_reorder(spreadsheet_id):
    print(f"--- Debugging Reorder for {spreadsheet_id} ---")
    ss = SheetsService()
    
    try:
        sh = ss.gc.open_by_key(spreadsheet_id)
        current_worksheets = sh.worksheets()
        current_names = [ws.title for ws in current_worksheets]
        print(f"Current Order: {', '.join(current_names)}")
        
        # Test the reordering logic
        result = ss.reorder_sheets(spreadsheet_id, SHEET_ORDER)
        
        print(f"Result Status: {result.get('message')}")
        print(f"Final Order: {', '.join(result.get('final_order', []))}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    # MT Source ID
    target_id = "1BW8Gk5_X2EZVjbnaa2yDm-bPzzlggwQrHepeNCcPCc0"
    debug_reorder(target_id)
