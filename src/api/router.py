"""
Main API router.
"""

from fastapi import APIRouter

from .webhooks import webhook_router
from .endpoints import api_router

from .products import router as products_router

router = APIRouter()

# Webhook endpoints (from Google Sheets)
router.include_router(webhook_router, prefix="/webhook", tags=["webhooks"])

# API endpoints
router.include_router(api_router, prefix="/api/v1", tags=["api"])
router.include_router(products_router, prefix="/api/v1/products", tags=["products"])
