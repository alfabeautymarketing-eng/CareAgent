"""
Main API router.
"""

from fastapi import APIRouter

from .endpoints import api_router
from .products import router as products_router
from .routes.supabase import router as supabase_router
from .webhooks import webhook_router

router = APIRouter()

# Webhook endpoints (from Google Sheets)
router.include_router(webhook_router, prefix="/webhook", tags=["webhooks"])

# API endpoints
router.include_router(api_router, prefix="/api/v1", tags=["api"])
router.include_router(products_router, prefix="/api/v1/products", tags=["products"])
router.include_router(supabase_router, prefix="/api/v1/supabase", tags=["supabase"])
