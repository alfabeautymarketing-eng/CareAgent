"""
Эндпоинты вебхуков для событий Google Таблиц.
"""

from datetime import datetime
from typing import Optional, List

from fastapi import APIRouter, HTTPException, Header, Request
from pydantic import BaseModel, Field

from src.utils.logger import logger
from src.utils.config import settings

webhook_router = APIRouter()


class SheetEvent(BaseModel):
    """Данные события из Google Таблицы."""
    event: str = Field(..., description="Тип события (onChange, onEdit)")
    project: str = Field(..., description="Код проекта (mt, sk, ss)")
    sheet: str = Field(..., description="Имя листа")
    range: Optional[str] = Field(None, description="Диапазон (например, 'A1')")
    changed_columns: Optional[List[str]] = Field(None, description="Список измененных колонок")
    user: Optional[str] = Field(None, description="Email пользователя")
    timestamp: Optional[datetime] = Field(None, description="Время события")

    class Config:
        title = "Событие Google Таблицы"


class TaskResponse(BaseModel):
    """Ответ с ID задачи для отслеживания."""
    task_id: str = Field(..., description="ID созданной задачи")
    status: str = Field(..., description="Статус принятия")
    message: str = Field(..., description="Сообщение")

    class Config:
        title = "Ответ о постановке задачи"


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


@webhook_router.post("/sheets/{project}", response_model=TaskResponse, summary="Обработка событий из Google Sheets")
async def handle_sheets_webhook(
    project: str,
    event: SheetEvent,
    request: Request,
    x_webhook_signature: Optional[str] = Header(None),
    x_webhook_timestamp: Optional[str] = Header(None),
):
    """
    Основной обработчик вебхуков из Google Sheets.
    Принимает события изменения данных (настройки, редактирование) 
    и ставит задачи на синхронизацию в очередь.
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


@webhook_router.post("/sync/{project}", response_model=TaskResponse, summary="Запуск полной синхронизации проекта")
async def trigger_manual_sync(project: str):
    """
    Запускает полную синхронизацию всех данных для конкретного проекта.
    Используйте этот эндпоинт, когда нужно принудительно обновить всё вручную.
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
