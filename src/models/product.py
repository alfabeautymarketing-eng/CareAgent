"""
Pydantic models for Product entity - the canonical server-side representation.

This is the "source of truth" for product data. Google Sheets reflects this data,
not the other way around.
"""
from pydantic import BaseModel, Field, computed_field
from typing import Optional, List, Dict, Any
from datetime import datetime
from enum import Enum

from .price_models import ProjectCode


class ProductStatus(str, Enum):
    """Product lifecycle status."""
    ACTIVE = "active"
    DISCONTINUED = "discontinued"
    PENDING = "pending"
    ARCHIVED = "archived"


class ProductType(str, Enum):
    """Product type classification."""
    MAIN = "main"          # Regular product
    TESTER = "tester"      # Tester version
    SAMPLES = "samples"    # Samples/probes
    PROMO = "promo"        # Promotional materials


class PriceInfo(BaseModel):
    """Price information for a product."""
    base_price: float = 0.0
    currency: str = "USD"
    last_updated: Optional[datetime] = None
    price_history: List[Dict[str, Any]] = Field(default_factory=list)


class SupplierInfo(BaseModel):
    """Supplier-specific product information."""
    article: str
    name_original: str
    barcode: Optional[str] = None
    units_per_pack: int = 1
    group: str = ""
    line: str = ""


class LocalizationInfo(BaseModel):
    """Localized product information."""
    name_ru: Optional[str] = None
    name_en: Optional[str] = None
    description_ru: Optional[str] = None
    description_en: Optional[str] = None


class Product(BaseModel):
    """
    Canonical Product entity.

    This is the server-side source of truth. Structure:
    - id: Unique identifier (e.g., "MT-123", "SK-456")
    - project: Project code (mt, sk, ss)
    - supplier: Original supplier data
    - localization: Translated names/descriptions
    - price: Current price information
    - metadata: Audit and version info
    """

    # Primary identification
    id: str = Field(..., description="Unique product ID (e.g., MT-123)")
    project: ProjectCode

    # Supplier data (from price list)
    supplier: SupplierInfo

    # Localization
    localization: LocalizationInfo = Field(default_factory=LocalizationInfo)

    # Price
    price: PriceInfo = Field(default_factory=PriceInfo)

    # Classification
    status: ProductStatus = ProductStatus.ACTIVE
    product_type: ProductType = ProductType.MAIN
    volume: str = ""

    # Relationships
    parent_id: Optional[str] = None  # For testers/samples linked to main product
    category_id: Optional[str] = None

    # Metadata
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)
    version: int = 1
    source: str = "price_import"  # Where this data came from

    @computed_field
    @property
    def display_name(self) -> str:
        """Get best available display name."""
        if self.localization.name_ru:
            return self.localization.name_ru
        if self.localization.name_en:
            return self.localization.name_en
        return self.supplier.name_original

    @computed_field
    @property
    def full_article(self) -> str:
        """Article with volume suffix if present."""
        if self.volume:
            return f"{self.supplier.article} {self.volume}"
        return self.supplier.article

    def to_sheets_row(self, columns: List[str]) -> List[Any]:
        """
        Convert product to a row for Google Sheets export.

        Args:
            columns: List of column names to include

        Returns:
            List of values in column order
        """
        flat = self._flatten()
        return [flat.get(col, "") for col in columns]

    def _flatten(self) -> Dict[str, Any]:
        """Flatten nested structure for sheets export."""
        return {
            "ID-P": self.id,
            "Артикул": self.supplier.article,
            "Название (англ)": self.supplier.name_original,
            "Название (рус)": self.localization.name_ru or "",
            "Объём": self.volume,
            "Штрих-код": self.supplier.barcode or "",
            "Кол-во в уп.": self.supplier.units_per_pack,
            "Группа": self.supplier.group,
            "Линейка": self.supplier.line,
            "Цена": self.price.base_price,
            "Статус": self.status.value,
            "Тип": self.product_type.value,
        }

    @classmethod
    def from_sheets_row(
        cls,
        row: Dict[str, Any],
        project: ProjectCode,
        column_mapping: Optional[Dict[str, str]] = None
    ) -> "Product":
        """
        Create Product from Google Sheets row data.

        Args:
            row: Dictionary of column -> value
            project: Project code
            column_mapping: Optional custom column name mapping

        Returns:
            Product instance
        """
        mapping = column_mapping or {
            "ID-P": "id",
            "Артикул": "article",
            "Название (англ)": "name_original",
            "Название (рус)": "name_ru",
            "Объём": "volume",
            "Штрих-код": "barcode",
            "Кол-во в уп.": "units_per_pack",
            "Группа": "group",
            "Линейка": "line",
            "Цена": "base_price",
        }

        # Extract values using mapping
        product_id = row.get("ID-P", row.get("id", ""))

        supplier = SupplierInfo(
            article=row.get("Артикул", row.get("article", "")),
            name_original=row.get("Название (англ)", row.get("name_original", "")),
            barcode=row.get("Штрих-код", row.get("barcode")),
            units_per_pack=int(row.get("Кол-во в уп.", row.get("units_per_pack", 1)) or 1),
            group=row.get("Группа", row.get("group", "")),
            line=row.get("Линейка", row.get("line", "")),
        )

        localization = LocalizationInfo(
            name_ru=row.get("Название (рус)", row.get("name_ru")),
            name_en=row.get("Название (англ)", row.get("name_original")),
        )

        price = PriceInfo(
            base_price=float(row.get("Цена", row.get("base_price", 0)) or 0),
        )

        return cls(
            id=product_id,
            project=project,
            supplier=supplier,
            localization=localization,
            price=price,
            volume=row.get("Объём", row.get("volume", "")),
        )


class ProductCollection(BaseModel):
    """Collection of products with metadata."""

    project: ProjectCode
    products: List[Product] = Field(default_factory=list)
    last_sync: Optional[datetime] = None
    sync_source: str = "sheets"

    def get_by_article(self, article: str) -> Optional[Product]:
        """Find product by supplier article."""
        for p in self.products:
            if p.supplier.article == article:
                return p
        return None

    def get_by_id(self, product_id: str) -> Optional[Product]:
        """Find product by ID."""
        for p in self.products:
            if p.id == product_id:
                return p
        return None

    def get_articles_map(self) -> Dict[str, str]:
        """Get article -> ID mapping."""
        return {p.supplier.article: p.id for p in self.products}
