import asyncio
import sys
import os

# Add project root to path
sys.path.append(os.getcwd())

from unittest.mock import MagicMock
sys.modules["supabase"] = MagicMock()
sys.modules["src.services.supabase_service"] = MagicMock()

from src.api.endpoints import get_menu_config_by_spreadsheet
from src.utils.logger import logger

# SK Main Spreadsheet ID (from endpoints.py)
SK_SPREADSHEET_ID = "1hSsS9_Iu_MgKWsoE19hAMouQInLGVFaBF6ZFG4Bsm1s"
# SS Main Spreadsheet ID (from endpoints.py)
SS_SPREADSHEET_ID = "1Bq2Pq0P1SQZfJNBZC3yduYCJmnyc4L4vmbLtvsVUkcg"

async def test_menu_config():
    print("Testing SK Menu Config...")
    try:
        config_sk = await get_menu_config_by_spreadsheet(SK_SPREADSHEET_ID)
        print("SK Config Retrieved Successfully")
        print(f"Project: {config_sk['project']}")
        print(f"Menu Title: {config_sk['menu_title']}")
        print(f"Menu Groups Count: {len(config_sk['menus'])}")
        # print(config_sk)
    except Exception as e:
        print(f"ERROR getting SK config: {e}")
        import traceback
        traceback.print_exc()

    print("\nTesting SS Menu Config...")
    try:
        config_ss = await get_menu_config_by_spreadsheet(SS_SPREADSHEET_ID)
        print("SS Config Retrieved Successfully")
        print(f"Project: {config_ss['project']}")
        print(f"Menu Title: {config_ss['menu_title']}")
        print(f"Menu Groups Count: {len(config_ss['menus'])}")
    except Exception as e:
        print(f"ERROR getting SS config: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_menu_config())
