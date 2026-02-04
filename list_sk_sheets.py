from src.services.sheets import SheetsService
from src.utils.config import settings
import json

def list_sheets():
    svc = SheetsService()
    sk_id = "1hSsS9_Iu_MgKWsoE19hAMouQInLGVFaBF6ZFG4Bsm1s"
    sh = svc.open_by_key(sk_id)
    worksheets = sh.worksheets()
    titles = [ws.title for ws in worksheets]
    print(json.dumps(titles, indent=4, ensure_ascii=False))

if __name__ == "__main__":
    list_sheets()
