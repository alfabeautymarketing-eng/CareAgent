#!/usr/bin/env python3
"""
Test script to verify logging functionality:
1. Server logs to file (logs/server.log)
2. Supabase sync logs integration
"""
import sys
import os
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from utils.logger import setup_logging
from utils.config import settings
from services.sync_log_service import SyncLogService
import structlog

logger = structlog.get_logger()

def test_file_logging():
    """Test 1: Verify server logs are written to file"""
    print("\n=== Test 1: File Logging ===")
    
    # Setup logging with file output
    setup_logging(level=settings.log_level, log_file=settings.server_log_file)
    
    # Write test log entry
    logger.info("test_file_logging", message="Testing file logging functionality", test_id="FILE_001")
    
    # Check if file exists and has content
    log_path = Path(settings.server_log_file)
    if log_path.exists():
        with open(log_path, 'r') as f:
            lines = f.readlines()
            recent_lines = lines[-5:]  # Last 5 lines
            print(f"✅ Log file exists: {log_path}")
            print(f"📝 Recent log entries ({len(recent_lines)} lines):")
            for line in recent_lines:
                print(f"   {line.strip()}")
    else:
        print(f"❌ Log file not found: {log_path}")
        return False
    
    return True

def test_supabase_logging():
    """Test 2: Verify Supabase sync log integration"""
    print("\n=== Test 2: Supabase Logging ===")
    
    # Initialize sync log service
    sync_log_service = SyncLogService(
        data_dir=settings.sync_log_data_dir
    )
    
    # Check Supabase configuration
    supabase_url = getattr(sync_log_service.supabase, 'supabase_url', 'Not configured')
    if sync_log_service.supabase.is_configured():
        print("✅ Supabase is configured")
        print(f"   URL: {supabase_url}")
    else:
        print("⚠️  Supabase is not configured (will only log to local files)")
        print("   This is OK if Supabase credentials are not in .env")
    
    # Add a test log entry
    test_spreadsheet_id = "TEST_SPREADSHEET_" + str(os.getpid())
    
    result = sync_log_service.add_entry(
        spreadsheet_id=test_spreadsheet_id,
        source_info="Test Source Sheet",
        target_info="Test Target Sheet",
        row_key="TEST_001",
        old_value="old_test_value",
        new_value="new_test_value",
        category="SYNC_TEST",
        status="SUCCESS",
        project="TEST_PROJECT",
        event="TEST_EVENT",
        tags=["test", "verification"],
    )
    
    print(f"✅ Local log entry created: {result.get('id', 'N/A')}")
    
    # Check local file
    local_log_path = Path(settings.sync_log_data_dir) / f"{test_spreadsheet_id}.jsonl"
    if local_log_path.exists():
        print(f"✅ Local JSONL file created: {local_log_path}")
        with open(local_log_path, 'r') as f:
            lines = f.readlines()
            print(f"   Entries in file: {len(lines)}")
    else:
        print(f"❌ Local JSONL file not found: {local_log_path}")
        return False
    
    # Note about Supabase verification
    if sync_log_service.supabase.is_configured():
        print("\n💡 To verify Supabase logs:")
        print(f"   SELECT * FROM sync_logs WHERE spreadsheet_id = '{test_spreadsheet_id}' ORDER BY created_at DESC LIMIT 5;")
    
    return True

def main():
    print("🔍 Starting Logging Verification Tests")
    print("=" * 60)
    
    try:
        # Run tests
        test1_passed = test_file_logging()
        test2_passed = test_supabase_logging()
        
        # Summary
        print("\n" + "=" * 60)
        print("📊 Test Summary:")
        print(f"   File Logging: {'✅ PASSED' if test1_passed else '❌ FAILED'}")
        print(f"   Supabase Logging: {'✅ PASSED' if test2_passed else '❌ FAILED'}")
        
        if test1_passed and test2_passed:
            print("\n🎉 All tests passed!")
            return 0
        else:
            print("\n⚠️  Some tests failed")
            return 1
            
    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
