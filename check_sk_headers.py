from src.services.sheets import SheetsService
from src.utils.config import settings
import json

def check_headers():
    svc = SheetsService()
    sk_id = "1hSsS9_Iu_MgKWsoE19hAMouQInLGVFaBF6ZFG4Bsm1s"
    
    sheets_to_check = ["Сертификация", "Заказ", "Заказ2026"]
    results = {}
    
    for title in sheets_to_check:
        try:
            ws = svc.get_worksheet(sk_id, title)
            headers = ws.row_values(1)
            results[title] = headers
        except Exception as e:
            results[title] = f"Error: {str(e)}"
            
    print(json.dumps(results, indent=4, ensure_ascii=False))

if __name__ == "__main__":
    check_headers()
