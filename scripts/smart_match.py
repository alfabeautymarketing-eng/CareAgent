#!/usr/bin/env python3
"""
Smart Product Matcher (CLI)
---------------------------
CLI wrapper around the ProductMatcher service.

Usage:
    python scripts/smart_match.py "Cream with Vitamin C 50ml"
"""

import sys
import json
from pathlib import Path

# Add project root
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

from src.services.product_matcher import ProductMatcher

# --- CONFIGURATION ---
SPREADSHEET_ID = "13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ" 
SHEET_NAME = "Сертификация"

def main():
    if len(sys.argv) < 2:
        print("Please provide a product name.")
        print('Example: python scripts/smart_match.py "Vitamin C Cream (Tester)"')
        return

    new_name = sys.argv[1]
    
    matcher = ProductMatcher(SPREADSHEET_ID, SHEET_NAME)
    
    print(f"⏳ Fetching data from sheet '{SHEET_NAME}'...")
    candidates = matcher.fetch_base_products()
    
    if not candidates:
        print("⚠️ No base products found. Check sheet/flags.")
        return

    print(f"🔍 Searching match for: '{new_name}'")
    match_result = matcher.find_best_match(new_name, candidates)
    
    if match_result:
        print("\n" + "="*50)
        print("RESULT:")
        print(json.dumps(match_result, indent=2, ensure_ascii=False))
        print("="*50)
        
        if match_result.get("match_found") and match_result.get("matched_product_details"):
            matched_product = match_result["matched_product_details"]
            print(f"\n✅ ACTION: Auto-fill new row with:")
            print(f"   Name ENG DS: {matched_product['name_eng_ds']}")
            print(f"   Name RUS DS: {matched_product['name_rus_ds']}")
    else:
        print("❌ No match found or error.")

if __name__ == "__main__":
    main()
