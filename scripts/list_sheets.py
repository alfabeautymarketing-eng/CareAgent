#!/usr/bin/env python3
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from src.services.sheets import SheetsService

PROJECT_MAP = {
    "MT": "199Np7xsBiBRQih5_tlUdpt6EmkfRGjZAhTvKm4Ua0Q6XEaMtvAmQUn0g",
     # "SS": "1sTgZa-n1aP7oIhyQfPeN8QDgDNnCubqMWAd-TKjKpJXWsQm_ZhXnojPD", # Checked one is enough usually, but let's check all
     # "SK": "1DJvK1vUT2OTubN0TLdZvsgYMSYByLHl8xTsus3K-KJ-VtJxgGnSw5Ih8"
}

def list_sheets():
    sheets = SheetsService()
    for name, ssid in PROJECT_MAP.items():
        print(f"--- Project {name} ---")
        try:
            sh = sheets.gc.open_by_key(ssid)
            for ws in sh.worksheets():
                print(f"  - '{ws.title}'")
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    list_sheets()
