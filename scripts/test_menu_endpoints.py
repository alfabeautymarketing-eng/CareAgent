#!/usr/bin/env python3
"""
Test script for menu API endpoints.

Tests:
1. GET /api/v1/menu/{project} - Get menu by project code
2. GET /api/v1/menu/config?spreadsheet_id=... - Get menu by spreadsheet ID (for GAS)
3. Validate menu structure (primary_menu, order_stages_menu, menu_groups)
"""
import requests
import json
import sys

BASE_URL = "http://localhost:8000"

# Actual spreadsheet IDs from the system
SPREADSHEET_IDS = {
    "MT": "13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ",  # MT Main
    "SK": "1hSsS9_Iu_MgKWsoE19hAMouQInLGVFaBF6ZFG4Bsm1s",  # SK Main
    "SS": "1Bq2Pq0P1SQZfJNBZC3yduYCJmnyc4L4vmbLtvsVUkcg",  # SS Main
}

PROJECTS = ["MT", "SK", "SS"]

def validate_menu_response(response_json):
    """Validate menu response structure"""
    required_fields = ["menu_title", "order_sheet", "sort_columns", "primary_menu", "order_stages_menu", "menu_groups"]

    for field in required_fields:
        if field not in response_json:
            print(f"   ❌ Missing field: {field}")
            return False

    # Check primary_menu structure
    primary = response_json.get("primary_menu", {})
    if "title" not in primary or "items" not in primary:
        print(f"   ❌ Invalid primary_menu structure")
        return False

    # Check order_stages_menu structure
    stages = response_json.get("order_stages_menu", {})
    if "title" not in stages or "items" not in stages:
        print(f"   ❌ Invalid order_stages_menu structure")
        return False

    # Check menu_groups
    groups = response_json.get("menu_groups", [])
    if not isinstance(groups, list) or len(groups) == 0:
        print(f"   ⚠️  Warning: menu_groups is empty or not a list")

    print(f"   ✅ Menu structure is valid")
    return True


def test_menu_by_project():
    """Test GET /api/v1/menu/{project} endpoint"""
    print("\n=== Test 1: Menu by Project (GET /api/v1/menu/{{project}}) ===")

    all_pass = True
    for project in PROJECTS:
        endpoint = f"{BASE_URL}/api/v1/menu/{project}"
        print(f"\n   Testing {project}:")
        print(f"   GET {endpoint}")

        try:
            response = requests.get(endpoint, timeout=10)
            print(f"   Status: {response.status_code}")

            if response.status_code == 200:
                result = response.json()
                print(f"   Menu title: {result.get('menu_title', 'N/A')}")
                print(f"   Primary menu items: {len(result.get('primary_menu', {}).get('items', []))}")
                print(f"   Order stages items: {len(result.get('order_stages_menu', {}).get('items', []))}")
                print(f"   Menu groups: {len(result.get('menu_groups', []))}")

                if validate_menu_response(result):
                    print(f"   ✅ {project} menu loaded successfully")
                else:
                    all_pass = False
            else:
                print(f"   ❌ Error: {response.status_code}")
                print(f"   {response.text[:200]}")
                all_pass = False

        except Exception as e:
            print(f"   ❌ Error: {e}")
            all_pass = False

    return all_pass


def test_menu_by_spreadsheet_id():
    """Test GET /api/v1/menu/config?spreadsheet_id=... endpoint (for GAS)"""
    print("\n=== Test 2: Menu by Spreadsheet ID (GET /api/v1/menu/config?spreadsheet_id=...) ===")

    all_pass = True
    for project, spreadsheet_id in SPREADSHEET_IDS.items():
        endpoint = f"{BASE_URL}/api/v1/menu/config"
        params = {"spreadsheet_id": spreadsheet_id}
        print(f"\n   Testing {project} (spreadsheet_id={spreadsheet_id[:20]}...):")
        print(f"   GET {endpoint}?spreadsheet_id={spreadsheet_id[:20]}...")

        try:
            response = requests.get(endpoint, params=params, timeout=10)
            print(f"   Status: {response.status_code}")

            if response.status_code == 200:
                result = response.json()
                print(f"   Menu title: {result.get('menu_title', 'N/A')}")
                print(f"   Primary menu items: {len(result.get('primary_menu', {}).get('items', []))}")
                print(f"   Order stages items: {len(result.get('order_stages_menu', {}).get('items', []))}")
                print(f"   Menu groups: {len(result.get('menu_groups', []))}")

                if validate_menu_response(result):
                    print(f"   ✅ {project} menu loaded successfully via spreadsheet_id")
                else:
                    all_pass = False
            else:
                print(f"   ❌ Error: {response.status_code}")
                print(f"   {response.text[:200]}")
                all_pass = False

        except Exception as e:
            print(f"   ❌ Error: {e}")
            all_pass = False

    return all_pass

def main():
    print("\n" + "=" * 70)
    print("🔍 TESTING UNIFIED MENU SYSTEM")
    print("=" * 70)

    test1_pass = test_menu_by_project()
    test2_pass = test_menu_by_spreadsheet_id()

    print("\n" + "=" * 70)
    print("📊 TEST SUMMARY:")
    print(f"   ✓ Menu by Project:       {'PASSED' if test1_pass else 'FAILED'}")
    print(f"   ✓ Menu by Spreadsheet:   {'PASSED' if test2_pass else 'FAILED'}")
    print("=" * 70)

    if test1_pass and test2_pass:
        print("\n✅ All menu endpoints working correctly!")
        print("   Ready to refresh menus in Google Sheets")
        return 0
    else:
        print("\n❌ Some tests failed. Check server logs.")
        return 1

if __name__ == "__main__":
    exit(main())
