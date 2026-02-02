#!/usr/bin/env python3
"""
Diagnostic script for sort endpoints.
Tests:
1. GET /api/v1/menu/sort-config
2. POST /api/v1/sort
"""
import requests
import sys

BASE_URL = "http://46.226.167.153:8000"
SS_ID = "13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ" # MT Main

def test_sort_config():
    print("\n--- Testing Sort Config ---")
    url = f"{BASE_URL}/api/v1/menu/sort-config?spreadsheet_id={SS_ID}"
    try:
        resp = requests.get(url, timeout=5)
        print(f"Status: {resp.status_code}")
        if resp.status_code == 200:
            config = resp.json()
            print(f"Config: {config}")
            return config
        else:
            print(f"Error: {resp.text}")
    except Exception as e:
        print(f"Failed: {e}")
    return None

def test_sort_execution(config):
    if not config:
        return
    
    print("\n--- Testing Sort Execution ---")
    url = f"{BASE_URL}/api/v1/sort"
    payload = {
        "spreadsheet_id": SS_ID,
        "sheet_name": config["order_sheet"],
        "column_name": config["sort_columns"]["manufacturer"],
        "ascending": True
    }
    
    print(f"Payload: {payload}")
    try:
        resp = requests.post(url, json=payload, timeout=30)
        print(f"Status: {resp.status_code}")
        print(f"Response: {resp.text}")
    except Exception as e:
        print(f"Failed: {e}")

if __name__ == "__main__":
    config = test_sort_config()
    test_sort_execution(config)
