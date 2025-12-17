"""
Webhook handlers for Google Sheets events.
"""

from datetime import datetime
from typing import Optional, List

from fastapi import APIRouter, HTTPException, Header, Request
from pydantic import BaseModel

from src.utils.logger import logger
from src.utils.config import settings

webhook_router = APIRouter()


class SheetEvent(BaseModel):
    """Google Sheets event payload."""

    event: str  # onChange, onEdit
    project: str  # mt, sk, ss
    sheet: str  # Sheet name
    range: Optional[str] = None
    changed_columns: Optional[List[str]] = None
    user: Optional[str] = None
    timestamp: Optional[datetime] = None


class TaskResponse(BaseModel):
    """Response with task ID for tracking."""

    task_id: str
    status: str
    message: str


def verify_webhook_signature(
    payload: bytes,
    signature: str,
    timestamp: str,
) -> bool:
    """
    Verify HMAC-SHA256 signature from Google Apps Script.

    Args:
        payload: Raw request body
        signature: X-Webhook-Signature header value
        timestamp: X-Webhook-Timestamp header value

    Returns:
        True if signature is valid
    """
    if not settings.webhook_secret:
        # No secret configured, skip verification (dev mode)
        return True

    import hmac
    import hashlib

    expected = hmac.new(
        settings.webhook_secret.encode(),
        payload,
        hashlib.sha256,
    ).hexdigest()

    provided = signature.replace("sha256=", "")
    return hmac.compare_digest(expected, provided)


@webhook_router.post("/sheets/{project}", response_model=TaskResponse)
async def handle_sheets_webhook(
    project: str,
    event: SheetEvent,
    request: Request,
    x_webhook_signature: Optional[str] = Header(None),
    x_webhook_timestamp: Optional[str] = Header(None),
):
    """
    Handle webhook from Google Sheets.

    This endpoint receives events when data changes in Google Sheets
    and queues sync tasks for processing.
    """
    # Validate project
    if project not in ["mt", "sk", "ss"]:
        raise HTTPException(status_code=404, detail=f"Unknown project: {project}")

    # Verify signature
    if x_webhook_signature:
        body = await request.body()
        if not verify_webhook_signature(body, x_webhook_signature, x_webhook_timestamp or ""):
            raise HTTPException(status_code=401, detail="Invalid webhook signature")

    # Log event
    logger.info(
        "webhook_received",
        project=project,
        event=event.event,
        sheet=event.sheet,
        range=event.range,
    )

    # TODO: Queue sync task
    # task_id = await sync_service.queue_sync(project, event)
    task_id = f"task_{datetime.now().strftime('%Y%m%d%H%M%S')}"

    return TaskResponse(
        task_id=task_id,
        status="accepted",
        message=f"Sync task queued for {project}/{event.sheet}",
    )


@webhook_router.post("/sync/{project}", response_model=TaskResponse)
async def trigger_manual_sync(project: str):
    """
    Manually trigger full sync for a project.

    Use this endpoint to force a complete sync when needed.
    """
    if project not in ["mt", "sk", "ss"]:
        raise HTTPException(status_code=404, detail=f"Unknown project: {project}")

    logger.info("manual_sync_triggered", project=project)

    # TODO: Queue full sync task
    task_id = f"fullsync_{datetime.now().strftime('%Y%m%d%H%M%S')}"

    return TaskResponse(
        task_id=task_id,
        status="accepted",
        message=f"Full sync queued for {project}",
    )
