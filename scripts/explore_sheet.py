import sys
import os
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

from src.services.sheets import SheetsService
from src.utils.config import settings

def main():
    if len(sys.argv) < 3:
        print("Usage: python scripts/explore_sheet.py <SPREADSHEET_ID> <SHEET_NAME>")
        sys.exit(1)

    spreadsheet_id = sys.argv[1]
    sheet_name = sys.argv[2]
    
    print(f"Connecting to {spreadsheet_id}...")
    svc = SheetsService()
    
    try:
        sh = svc.gc.open_by_key(spreadsheet_id)
        print(f"Spreadsheet Title: {sh.title}")
        
        if len(sys.argv) > 2:
            sheet_name = sys.argv[2]
            print(f"Fetching headers for '{sheet_name}'...")
            try:
                headers = svc.get_worksheet_headers(spreadsheet_id, sheet_name)
                print("\nFOUND HEADERS:")
                for i, h in enumerate(headers):
                    print(f"{i+1}: {h}")
            except Exception as e:
                print(f"❌ Error fetching headers for '{sheet_name}': {e}")
                print("\nAvailable sheets:")
                for ws in sh.worksheets():
                    print(f" - {ws.title}")
        else:
            print("\nAvailable sheets:")
            for ws in sh.worksheets():
                print(f" - {ws.title}")
                
    except Exception as e:
        print(f"❌ critical error: {e}")

if __name__ == "__main__":
    main()
