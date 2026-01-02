"""
Concrete store implementations for key entities.

These stores provide typed access to entity data with full CRUD operations.
Data is stored in JSON files and can be synced with Google Sheets.
"""
from typing import Optional, List, Dict, Any
from datetime import datetime

from .json_store import JSONStore
from src.models import (
    Product,
    ProductCollection,
    Certification,
    CertificationCollection,
    Order,
    OrderCollection,
    ProjectCode,
    OrderStage,
)
from src.utils.logger import logger


class ProductStore(JSONStore[Product]):
    """Store for Product entities."""

    def __init__(self, data_dir: str = "data"):
        super().__init__(
            model_class=Product,
            entity_name="products",
            data_dir=data_dir,
            id_field="id"
        )

    def get_by_article(self, article: str, project: Optional[str] = None) -> Optional[Product]:
        """Find product by supplier article."""
        entities = self.get_all(project)
        for p in entities:
            if p.supplier.article == article:
                return p
        return None

    def get_collection(self, project: str) -> ProductCollection:
        """Get all products as a collection."""
        products = self.get_all(project)
        return ProductCollection(
            project=ProjectCode(project),
            products=products,
            last_sync=datetime.utcnow()
        )

    def get_articles_map(self, project: str) -> Dict[str, str]:
        """Get article -> ID mapping for a project."""
        entities = self.get_all(project)
        return {p.supplier.article: p.id for p in entities}

    def bulk_upsert_from_price(
        self,
        processed_rows: List[Dict[str, Any]],
        project: str
    ) -> Dict[str, int]:
        """
        Bulk upsert products from price processing results.

        Returns:
            Dict with counts: created, updated, skipped
        """
        stats = {"created": 0, "updated": 0, "skipped": 0}
        existing = self.get_articles_map(project)

        for row in processed_rows:
            article = row.get("article", "")
            if not article:
                stats["skipped"] += 1
                continue

            try:
                if article in existing:
                    # Update existing product
                    product = self.get(existing[article], project)
                    if product:
                        product.price.base_price = float(row.get("price", 0))
                        product.supplier.barcode = row.get("barcode", product.supplier.barcode)
                        product.updated_at = datetime.utcnow()
                        self.save(product)
                        stats["updated"] += 1
                else:
                    # Create new product
                    product = Product.from_sheets_row(row, ProjectCode(project))
                    self.save(product)
                    stats["created"] += 1
            except Exception as e:
                logger.warning(f"Failed to upsert product {article}: {e}")
                stats["skipped"] += 1

        return stats


class CertificationStore(JSONStore[Certification]):
    """Store for Certification entities."""

    def __init__(self, data_dir: str = "data"):
        super().__init__(
            model_class=Certification,
            entity_name="certifications",
            data_dir=data_dir,
            id_field="id"
        )

    def get_by_document_number(
        self,
        doc_number: str,
        project: Optional[str] = None
    ) -> Optional[Certification]:
        """Find certification by document number."""
        entities = self.get_all(project)
        for c in entities:
            if c.document_number == doc_number:
                return c
        return None

    def get_for_product(self, product_id: str, project: Optional[str] = None) -> List[Certification]:
        """Get all certifications covering a specific product."""
        entities = self.get_all(project)
        return [c for c in entities if product_id in c.product_ids]

    def get_expiring_soon(self, days: int = 30, project: Optional[str] = None) -> List[Certification]:
        """Get certifications expiring within specified days."""
        entities = self.get_all(project)
        result = []
        for c in entities:
            if c.days_until_expiry is not None and 0 < c.days_until_expiry <= days:
                result.append(c)
        return result

    def get_expired(self, project: Optional[str] = None) -> List[Certification]:
        """Get all expired certifications."""
        entities = self.get_all(project)
        return [c for c in entities if not c.is_valid]

    def get_collection(self, project: str) -> CertificationCollection:
        """Get all certifications as a collection."""
        certs = self.get_all(project)
        return CertificationCollection(
            project=ProjectCode(project),
            certifications=certs,
            last_sync=datetime.utcnow()
        )


class OrderStore(JSONStore[Order]):
    """Store for Order entities."""

    def __init__(self, data_dir: str = "data"):
        super().__init__(
            model_class=Order,
            entity_name="orders",
            data_dir=data_dir,
            id_field="id"
        )

    def get_by_order_number(
        self,
        order_number: str,
        project: Optional[str] = None
    ) -> Optional[Order]:
        """Find order by supplier order number."""
        entities = self.get_all(project)
        for o in entities:
            if o.order_number == order_number:
                return o
        return None

    def get_by_stage(self, stage: OrderStage, project: Optional[str] = None) -> List[Order]:
        """Get all orders at a specific stage."""
        entities = self.get_all(project)
        return [o for o in entities if o.stage == stage]

    def get_active(self, project: Optional[str] = None) -> List[Order]:
        """Get all active orders."""
        entities = self.get_all(project)
        return [o for o in entities if o.is_active]

    def get_for_product(self, product_id: str, project: Optional[str] = None) -> List[Order]:
        """Get all orders containing a specific product."""
        entities = self.get_all(project)
        result = []
        for order in entities:
            if any(item.product_id == product_id for item in order.items):
                result.append(order)
        return result

    def transition_order(
        self,
        order_id: str,
        new_stage: OrderStage,
        notes: str = "",
        changed_by: str = None,
        project: Optional[str] = None
    ) -> Optional[Order]:
        """Transition an order to a new stage."""
        order = self.get(order_id, project)
        if order:
            order.transition_to(new_stage, notes, changed_by)
            self.save(order)
            return order
        return None

    def get_collection(self, project: str) -> OrderCollection:
        """Get all orders as a collection."""
        orders = self.get_all(project)
        return OrderCollection(
            project=ProjectCode(project),
            orders=orders,
            last_sync=datetime.utcnow()
        )

    def generate_order_id(self, project: str) -> str:
        """Generate a unique order ID for the project."""
        year = datetime.utcnow().year
        prefix = f"ORD-{project.upper()}-{year}"

        # Find highest existing number
        existing = self.get_all(project)
        max_num = 0

        for order in existing:
            if order.id.startswith(prefix):
                try:
                    num = int(order.id.split("-")[-1])
                    if num > max_num:
                        max_num = num
                except ValueError:
                    pass

        return f"{prefix}-{max_num + 1:03d}"


# Singleton instances for convenience
_product_store: Optional[ProductStore] = None
_certification_store: Optional[CertificationStore] = None
_order_store: Optional[OrderStore] = None


def get_product_store(data_dir: str = "data") -> ProductStore:
    """Get or create ProductStore singleton."""
    global _product_store
    if _product_store is None:
        _product_store = ProductStore(data_dir)
    return _product_store


def get_certification_store(data_dir: str = "data") -> CertificationStore:
    """Get or create CertificationStore singleton."""
    global _certification_store
    if _certification_store is None:
        _certification_store = CertificationStore(data_dir)
    return _certification_store


def get_order_store(data_dir: str = "data") -> OrderStore:
    """Get or create OrderStore singleton."""
    global _order_store
    if _order_store is None:
        _order_store = OrderStore(data_dir)
    return _order_store
