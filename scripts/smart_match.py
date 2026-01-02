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

import argparse

def main():
    parser = argparse.ArgumentParser(description="Smart Product Matcher using Gemini AI.")
    parser.add_argument("query", help="Product name to match")
    parser.add_argument("-v", "--verbose", action="store_true", help="Show detailed matching process")
    args = parser.parse_args()

    new_name = args.query
    
    matcher = ProductMatcher(SPREADSHEET_ID, SHEET_NAME)
    
    print(f"⏳ Fetching data from sheet '{SHEET_NAME}'...")
    candidates = matcher.fetch_base_products()
    
    if not candidates:
        print("⚠️ No base products found. Check sheet/flags.")
        return

    print(f"🔍 Searching match for: '{new_name}'")
    if args.verbose:
        print(f"📦 Comparing against {len(candidates)} base products:")

    # 2. Match
    match_result = matcher.find_best_match(new_name, candidates)
    
    # 3. Output
    if match_result:
        if args.verbose:
            print("\n" + "="*50)
            print("FULL AI RESPONSE:")
            print(json.dumps(match_result, indent=2, ensure_ascii=False))
            print("="*50)
            
        if match_result.get("match_found"):
            matched_product = match_result.get("matched_product_details")
            if matched_product:
                print(f"\n✅ MATCH FOUND ({match_result['confidence']}%):")
                print(f"   Reason: {match_result['reasoning']}")
                print(f"\n👉 SUGGESTED AUTO-FILL:")
                print(f"   Name ENG DS: {matched_product['name_eng_ds']}")
                print(f"   Name RUS DS: {matched_product['name_rus_ds']}")
        else:
            print("\n❌ No confident match found.")
            if not args.verbose:
                print(f"   Reason: {match_result.get('reasoning', 'No reason provided')}")
                print("   (Try running with -v for more details)")
    else:
        print("❌ Error during AI matching process.")

if __name__ == "__main__":
    main()
