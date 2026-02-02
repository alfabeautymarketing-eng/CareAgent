
import gspread
from google.oauth2.service_account import Credentials
import yaml
import os

# Project IDs from yaml configs
spreadsheets = {
    "MT": "13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ",
    "SK": "1hSsS9_Iu_MgKWsoE19hAMouQInLGVFaBF6ZFG4Bsm1s",
    "SS": "1Bq2Pq0P1SQZfJNBZC3yduYCJmnyc4L4vmbLtvsVUkcg"
}

creds_path = "/Users/aleksandr/Desktop/AgentCare/config/credentials.json"
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def check_headers():
    if not os.path.exists(creds_path):
        print(f"Error: {creds_path} not found")
        return

    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)

    for name, sid in spreadsheets.items():
        try:
            ss = gc.open_by_key(sid)
            ws = ss.worksheet("Главная")
            headers = ws.row_values(1)
            print(f"Project {name} ({sid}):")
            print(f"Headers: {headers}")
            print(f"Count: {len(headers)}")
            print("-" * 20)
        except Exception as e:
            print(f"Failed to check {name}: {e}")

if __name__ == "__main__":
    check_headers()
