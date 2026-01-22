#!/usr/bin/env python3
"""
Test script for menu functions:
1. serverProcessPrimaryData -> /api/v1/price/process/{project}
2. serverLoadStockData -> /api/v1/stocks/load
"""
import requests
import json

BASE_URL = "http://localhost:8000"

def test_price_process_endpoint():
    """Test price processing endpoint"""
    print("\n=== Test 1: Price Processing Endpoint ===")
    
    # Test with MT project
    endpoint = f"{BASE_URL}/api/v1/price/process/mt"
    payload = {
        "spreadsheet_id": "test_spreadsheet_12345",
        "mode": "main",
        "dry_run": True  # Preview only, don't write
    }
    
    print(f"POST {endpoint}")
    print(f"Payload: {json.dumps(payload, indent=2)}")
    
    try:
        response = requests.post(endpoint, json=payload, timeout=10)
        print(f"\nStatus Code: {response.status_code}")
        
        if response.status_code in [200, 404, 422]:
            result = response.json()
            print(f"Response: {json.dumps(result, indent=2, ensure_ascii=False)}")
            
            if response.status_code == 200:
                print("✅ Endpoint exists and responds")
                return True
            elif response.status_code == 404:
                print("❌ Endpoint NOT FOUND - needs to be implemented")
                return False
            else:
                print("⚠️  Endpoint exists but has validation errors")
                return True
        else:
            print(f"Response text: {response.text[:500]}")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def test_stock_load_endpoint():
    """Test stock loading endpoint"""
    print("\n=== Test 2: Stock Loading Endpoint ===")
    
    endpoint = f"{BASE_URL}/api/v1/stocks/load"
    payload = {
        "spreadsheet_id": "test_spreadsheet_12345",
        "source_doc_id": None
    }
    
    print(f"POST {endpoint}")
    print(f"Payload: {json.dumps(payload, indent=2)}")
    
    try:
        response = requests.post(endpoint, json=payload, timeout=10)
        print(f"\nStatus Code: {response.status_code}")
        
        if response.status_code in [200, 404, 422]:
            result = response.json()
            print(f"Response: {json.dumps(result, indent=2, ensure_ascii=False)}")
            
            if response.status_code == 200:
                print("✅ Endpoint exists and responds")
                return True
            elif response.status_code == 404:
                print("❌ Endpoint NOT FOUND - needs to be implemented")
                return False
            else:
                print("⚠️  Endpoint exists but has validation errors")
                return True
        else:
            print(f"Response text: {response.text[:500]}")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def main():
    print("🔍 Testing Menu Function Endpoints")
    print("=" * 60)
    
    test1_pass = test_price_process_endpoint()
    test2_pass = test_stock_load_endpoint()
    
    print("\n" + "=" * 60)
    print("📊 Test Summary:")
    print(f"   Price Processing: {'✅ EXISTS' if test1_pass else '❌ MISSING'}")
    print(f"   Stock Loading: {'✅ EXISTS' if test2_pass else '❌ MISSING'}")
    
    if test1_pass and test2_pass:
        print("\n🎉 Both endpoints are available!")
    else:
        print("\n⚠️  Some endpoints need to be implemented")
    
    return 0 if (test1_pass and test2_pass) else 1

if __name__ == "__main__":
    exit(main())
