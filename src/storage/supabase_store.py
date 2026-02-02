from typing import List, Optional, Dict, Any
from datetime import datetime
import uuid

from src.models.product import Product, ProductCollection, ProjectCode, SupplierInfo, LocalizationInfo, PriceInfo, ProductStatus, ProductType
from src.services.supabase_service import get_supabase_service
from src.utils.logger import logger

class SupabaseProductStore:
    """
    Store for Product entities backed by Supabase 'articles' table.
    """

    def __init__(self):
        self.supabase = get_supabase_service()
        self.table_name = "articles"

    def _map_to_db(self, product: Product) -> Dict[str, Any]:
        """Convert Product entity to Supabase table row."""
        return {
            "project_key": product.project.value,
            "article_id_p": product.id,
            "article_id_g": product.id_g,
            "article_id_l": product.id_l,
            "article_id": product.supplier.article,
            "name_eng_price": product.supplier.name_original,
            "article_rus": product.localization.name_ru,
            "volume": product.volume,
            "units_per_pack": str(product.supplier.units_per_pack),
            "group_name": product.supplier.group,
            "line_name": product.supplier.line,
            "barcode": product.supplier.barcode,
            "price": product.price.base_price,
            "status": product.status.value,
            "product_type": product.product_type.value,
            "updated_at": product.updated_at.isoformat() if product.updated_at else datetime.utcnow().isoformat(),
            # "created_at" is handled by Supabase default if new
        }

    def _map_from_db(self, row: Dict[str, Any]) -> Product:
        """Convert Supabase table row to Product entity."""
        project = ProjectCode(row.get("project_key", "mt")) # Default to MT if missing
        
        supplier = SupplierInfo(
            article=row.get("article_id", ""),
            name_original=row.get("name_eng_price", ""),
            barcode=row.get("barcode"),
            units_per_pack=int(row.get("units_per_pack", "1") or 1),
            group=row.get("group_name", ""),
            line=row.get("line_name", ""),
        )

        localization = LocalizationInfo(
            name_ru=row.get("article_rus"),
        )

        price = PriceInfo(
            base_price=float(row.get("price") or 0.0),
        )
        
        # Handle enums safely
        try:
            status = ProductStatus(row.get("status", "active"))
        except ValueError:
            status = ProductStatus.ACTIVE

        try:
            p_type = ProductType(row.get("product_type", "main"))
        except ValueError:
            p_type = ProductType.MAIN

        return Product(
            id=row.get("article_id_p", ""),
            id_g=row.get("article_id_g"),
            id_l=row.get("article_id_l"),
            project=project,
            supplier=supplier,
            localization=localization,
            price=price,
            volume=row.get("volume", ""),
            status=status,
            product_type=p_type,
            updated_at=datetime.fromisoformat(row.get("updated_at")) if row.get("updated_at") else datetime.utcnow(),
            source="supabase"
        )

    def save(self, product: Product) -> bool:
        """Save (upsert) product to Supabase."""
        data = self._map_to_db(product)
        # We need to handle the PK. 
        # If we use article_id_p as a unique key for upsert, we might need a unique index on it in DB.
        # Alternatively, we can look up by article_id_p first.
        # For now, let's try upserting with conflict on article_id_p if possible.
        # BUT, the table PK is likely 'id' (uuid). 
        # If we don't provide 'id', Supabase creates a new one.
        # To update, we need to know the 'id' OR have a unique constraint on 'article_id_p'.
        # Assuming we might not have 'id' in Product yet.
        
        # Strategy: upsert based on 'article_id_p' IF it has a unique constraint.
        # If not, we run a select first.
        try:
            # Try to find existing record by article_id_p and project_key
            existing = self.supabase.client.table(self.table_name)\
                .select("id")\
                .eq("article_id_p", product.id)\
                .eq("project_key", product.project.value)\
                .execute()
            
            if existing.data and len(existing.data) > 0:
                # Update
                record_id = existing.data[0]['id']
                self.supabase.client.table(self.table_name)\
                    .update(data)\
                    .eq("id", record_id)\
                    .execute()
            else:
                # Insert
                self.supabase.client.table(self.table_name).insert(data).execute()
            
            return True
        except Exception as e:
            logger.error("supabase_save_failed", product_id=product.id, error=str(e))
            return False

    def bulk_upsert(self, products: List[Product]) -> Dict[str, int]:
        """Bulk upsert products."""
        # This is tricky without a unique constraint we can rely on for ON CONFLICT.
        # For efficiency, we should probably rely on a unique constraint on (project_key, article_id_p).
        # We will assume for now we do one by one or small batches, OR we implement the check logic.
        
        # Better approach for bulk:
        # 1. Get all existing IDs for the project
        # 2. Separate into inserts and updates
        
        stats = {"created": 0, "updated": 0, "errors": 0}
        
        if not products:
            return stats
            
        project_key = products[0].project.value # Assume all same project for now
        
        try:
            # 1. Fetch all existing article_id_p for this project
            # Optimization: only fetch those we are trying to upsert?
            # For large sets, mapping one by one is slow.
            
            for p in products:
                if self.save(p):
                    # We can't easily distinguish create/update with the simple save logic above without checking response
                    # But save() handles it.
                    stats["updated"] += 1 # We'll just count successes
                else:
                    stats["errors"] += 1
                    
            return stats
        except Exception as e:
            logger.error("supabase_bulk_upsert_failed", error=str(e))
            return stats

    def get_by_id(self, product_id: str, project_key: str) -> Optional[Product]:
        """Get product by ID-P."""
        try:
            response = self.supabase.client.table(self.table_name)\
                .select("*")\
                .eq("article_id_p", product_id)\
                .eq("project_key", project_key)\
                .execute()
            
            if response.data and len(response.data) > 0:
                return self._map_from_db(response.data[0])
            return None
        except Exception as e:
            logger.error("supabase_get_failed", product_id=product_id, error=str(e))
            return None

    def get_all(self, project_key: str) -> List[Product]:
        """Get all products for a project."""
        try:
            # Supabase default limit is 1000. Use pagination if needed.
            # implementing simple fetch for now.
            products = []
            start = 0
            batch_size = 1000
            
            while True:
                response = self.supabase.client.table(self.table_name)\
                    .select("*")\
                    .eq("project_key", project_key)\
                    .range(start, start + batch_size - 1)\
                    .execute()
                
                if not response.data:
                    break
                    
                for row in response.data:
                    products.append(self._map_from_db(row))
                    
                if len(response.data) < batch_size:
                    break
                    
                start += batch_size
                
            return products
        except Exception as e:
            logger.error("supabase_get_all_failed", error=str(e))
            return []

_supabase_product_store = None

def get_supabase_product_store() -> SupabaseProductStore:
    global _supabase_product_store
    if not _supabase_product_store:
        _supabase_product_store = SupabaseProductStore()
    return _supabase_product_store
