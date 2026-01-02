"""
Storage module for server-first data management.

This module provides abstractions for storing canonical data on the server,
with Google Sheets serving as the view/interface layer.
"""

from .base import BaseStore, StoreError
from .json_store import JSONStore
from .stores import (
    ProductStore,
    CertificationStore,
    OrderStore,
    get_product_store,
    get_certification_store,
    get_order_store,
)

__all__ = [
    "BaseStore",
    "StoreError",
    "JSONStore",
    "ProductStore",
    "CertificationStore",
    "OrderStore",
    "get_product_store",
    "get_certification_store",
    "get_order_store",
]
