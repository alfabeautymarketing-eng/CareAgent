"""
Data models for AgentCare server-first architecture.

This module contains canonical Pydantic models that serve as the
"source of truth" for all entities. Google Sheets reflects this data.
"""

from .price_models import (
    ProjectCode,
    ProcessingMode,
    PriceProcessRequest,
    ProcessedRow,
    ProcessedData,
    SyncResult,
    PriceProcessResponse,
    StockLoadRequest,
    StockLoadResponse,
)

from .product import (
    ProductStatus,
    ProductType,
    PriceInfo,
    SupplierInfo,
    LocalizationInfo,
    Product,
    ProductCollection,
)

from .certification import (
    CertificationType,
    CertificationStatus,
    CertificationDocument,
    ProductReference,
    Certification,
    CertificationCollection,
)

from .order import (
    OrderStage,
    OrderItemStatus,
    OrderItem,
    StageTransition,
    ShippingInfo,
    Order,
    OrderCollection,
)

__all__ = [
    # Price models
    "ProjectCode",
    "ProcessingMode",
    "PriceProcessRequest",
    "ProcessedRow",
    "ProcessedData",
    "SyncResult",
    "PriceProcessResponse",
    "StockLoadRequest",
    "StockLoadResponse",
    # Product models
    "ProductStatus",
    "ProductType",
    "PriceInfo",
    "SupplierInfo",
    "LocalizationInfo",
    "Product",
    "ProductCollection",
    # Certification models
    "CertificationType",
    "CertificationStatus",
    "CertificationDocument",
    "ProductReference",
    "Certification",
    "CertificationCollection",
    # Order models
    "OrderStage",
    "OrderItemStatus",
    "OrderItem",
    "StageTransition",
    "ShippingInfo",
    "Order",
    "OrderCollection",
]
