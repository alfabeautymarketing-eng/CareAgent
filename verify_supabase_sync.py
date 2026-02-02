import asyncio
import os
from src.storage.supabase_store import get_supabase_product_store
from src.models.product import Product, ProjectCode, SupplierInfo, LocalizationInfo, PriceInfo, ProductStatus, ProductType
from src.utils.logger import logger

async def test_supabase_integration():
    print("🚀 Starting Supabase Integration Test...")
    
    store = get_supabase_product_store()
    
    # 1. Create a test product
    test_id = "TEST-SUPABASE-001"
    test_product = Product(
        id=test_id,
        id_g="TEST-G-001",
        id_l="TEST-L-001",
        project=ProjectCode.MT,
        supplier=SupplierInfo(
            article=test_id,
            name_original="Test Product for Supabase",
            group="Test Group",
            line="Test Line",
            barcode="123456789",
            units_per_pack=1
        ),
        localization=LocalizationInfo(
            name_ru="Тестовый Товар",
            name_en="Test Product"
        ),
        price=PriceInfo(
            base_price=99.99
        ),
        volume="100ml",
        status=ProductStatus.ACTIVE,
        product_type=ProductType.MAIN
    )
    
    try:
        # 2. Save (Upsert)
        print(f"📦 Saving product {test_id} to Supabase...")
        stats = store.bulk_upsert([test_product])
        print(f"✅ Upsert stats: {stats}")
        
        if stats.get('errors', 0) > 0:
            print("❌ Errors during upsert!")
            return

        # 3. Read back to verify
        print("🔍 Verifying data in Supabase...")
        
        res = store.supabase.client.table("articles").select("*").eq("article_id_p", test_id).execute()
        if res.data:
            item = res.data[0]
            print(f"✅ Product found in DB:")
            print(f"   - ID: {item.get('article_id_p')}")
            print(f"   - Name RU: {item.get('article_rus')}")
            print(f"   - Price: {item.get('price')}")
            print(f"   - Type: {item.get('product_type')}")
            
            assert item.get('price') == 99.99, f"Expected price 99.99, got {item.get('price')}"
            assert item.get('product_type') == "main", f"Expected type 'main', got {item.get('product_type')}"
            print("✅ Data validation passed!")
            
            # 4. Clean up
            print("🧹 Cleaning up test data...")
            store.supabase.client.table("articles").delete().eq("article_id_p", test_id).execute()
            print("✅ Test data deleted.")
            
        else:
            print("❌ Product NOT found in DB after upsert!")

    except Exception as e:
        print(f"❌ Test Failed with Exception: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_supabase_integration())
