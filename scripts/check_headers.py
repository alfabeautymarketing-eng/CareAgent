import os
import json
import sys
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def get_headers(spreadsheet_id, range_name):
    # Path to service account key
    creds_path = '/Users/aleksandr/Desktop/AgentCare/config/credentials.json'
    if not os.path.exists(creds_path):
        print(f"Error: {creds_path} not found")
        return None
    
    creds = Credentials.from_service_account_file(creds_path, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
    service = build('sheets', 'v4', credentials=creds)
    
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    if values:
        return values[0]
    return []

spreadsheet_id = "1hSsS9_Iu_MgKWsoE19hAMouQInLGVFaBF6ZFG4Bsm1s" # SK (Carmado)
sheets = ["Главная", "Сертификация", "Заказ", "Заказ2026", "Этикетки", "Прайс", "Динамика цены", "Расчет цены"]

for sheet in sheets:
    print(f"\n--- Headers for {sheet} ---")
    headers = get_headers(spreadsheet_id, f"'{sheet}'!1:1")
    if headers:
        for i, h in enumerate(headers):
            print(f"{i}: {h}")
    else:
        print("No headers found or error")
