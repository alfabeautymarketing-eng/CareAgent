"""
Main API router.
"""

from fastapi import APIRouter

from .webhooks import webhook_router
from .endpoints import api_router

router = APIRouter()

# Webhook endpoints (from Google Sheets)
router.include_router(webhook_router, prefix="/webhook", tags=["webhooks"])

# API endpoints
router.include_router(api_router, prefix="/api/v1", tags=["api"])
