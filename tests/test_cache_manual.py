import sys
import os
from pathlib import Path

# Add src to sys.path
sys.path.append(str(Path(__file__).parent.parent))

from src.utils.cache import CacheService
import time

def test_cache():
    print("Testing CacheService...")
    cache = CacheService()
    
    # Test set/get
    cache.set("test_key", {"a": 1}, ttl=2)
    val = cache.get("test_key")
    assert val == {"a": 1}, f"Expected {{'a': 1}}, got {val}"
    print("✓ Set/Get works")
    
    # Test expiry
    print("Waiting for expiry (2s)...")
    time.sleep(2.1)
    val = cache.get("test_key")
    assert val is None, f"Expected None after expiry, got {val}"
    print("✓ Expiry works")
    
    # Test delete
    cache.set("delete_key", "data")
    cache.delete("delete_key")
    val = cache.get("delete_key")
    assert val is None, "Expected None after delete"
    print("✓ Delete works")
    
    print("All CacheService tests passed!")

if __name__ == "__main__":
    try:
        test_cache()
    except Exception as e:
        print(f"Test failed: {e}")
        sys.exit(1)
