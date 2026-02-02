
import os
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

def check_headers():
    load_dotenv()
    
    # Path to service account key
    creds_path = os.getenv("GOOGLE_CREDENTIALS_FILE")
    if not creds_path:
        print("GOOGLE_CREDENTIALS_FILE not set")
        return

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)
    
    # SS spreadsheet ID from 01Config.js
    spreadsheet_id = "1Bq2Pq0P1SQZfJNBZC3yduYCJmnyc4L4vmbLtvsVUkcg"
    
    try:
        ss = gc.open_by_key(spreadsheet_id)
        ws = ss.worksheet("Заказ")
        headers = ws.row_values(1)
        
        print(f"Headers for sheet 'Заказ':")
        for i, h in enumerate(headers):
            print(f"  [{i}] '{h}' (len: {len(h)}, spaces: {h.count(' ')})")
            if "Название" in h and "ENG" in h:
                # Show exact representation
                print(f"    EXACT: {repr(h)}")
                
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_headers()
