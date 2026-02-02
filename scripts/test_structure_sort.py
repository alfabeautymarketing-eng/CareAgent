
import requests
import json

BASE_URL = "http://46.226.167.153:8000"
SS_ID = "13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ" # MT Main

def test_structure_sort():
    print("\n--- Testing Structural Sort (byManufacturer) ---")
    url = f"{BASE_URL}/api/v1/sort/structure"
    payload = {
        "spreadsheet_id": SS_ID,
        "mode": "byManufacturer"
    }
    
    print(f"Payload: {payload}")
    try:
        resp = requests.post(url, json=payload, timeout=60)
        print(f"Status: {resp.status_code}")
        if resp.status_code == 200:
            print(f"Success: {json.dumps(resp.json(), indent=2, ensure_ascii=False)}")
        else:
            print(f"Error: {resp.text}")
    except Exception as e:
        print(f"Failed: {e}")

if __name__ == "__main__":
    test_structure_sort()
