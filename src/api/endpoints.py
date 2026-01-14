"""
API endpoints for sync, price processing, AI analysis.
"""

from datetime import datetime
from typing import Optional, List, Any, Dict

from fastapi import APIRouter, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel

from src.utils.logger import logger
from src.utils.config import settings

api_router = APIRouter()

from src.services.sheets import SheetsService
from src.services.sync import SyncService
from src.services.sorting import SortingService
from src.services.logging import LoggingService
from src.services.ai import AIService, get_ai_service
from src.services.price_processor import PriceProcessor, get_price_processor
from src.services.invoice_service import get_invoice_service
from src.services.external_docs import ExternalDocsService
from src.services.sync_log_service import SyncLogService
from src.services.function_log_service import FunctionLogService
from src.services.certification import CertificationService
from src.services.meta_cache import MetaCacheService

sheets_service = SheetsService()
function_log_service = FunctionLogService()
logging_service = LoggingService(sheets_service)
sync_log_service = SyncLogService(
    data_dir=settings.sync_log_data_dir,
    retention_days=settings.sync_log_retention_days,
    max_entries=settings.sync_log_max_entries,
)
sync_service = SyncService(logging_service, sync_log_service) # Pass logger to sync service
sorting_service = SortingService()
ai_service = get_ai_service()
price_processor = get_price_processor(sheets_service, sync_service, logging_service)
external_docs_service = ExternalDocsService()
meta_cache_service = MetaCacheService()
certification_service = CertificationService(sheets_service)

class RuleItem(BaseModel):
    mode: str = "unidirectional"
    enabled: bool = True
    category: str = ""
    source_sheet: Optional[str] = None
    source_header: Optional[str] = None
    target_sheet: Optional[str] = None
    target_header: Optional[str] = None
    sheet_a: Optional[str] = None
    header_a: Optional[str] = None
    sheet_b: Optional[str] = None
    header_b: Optional[str] = None
    is_external: bool = False
    target_doc_id: Optional[str] = None

class RulesSaveRequest(BaseModel):
    rules: List[RuleItem]


# CRUD Models for Sync Rules
class RuleCreateRequest(BaseModel):
    """Request to create a new sync rule."""
    mode: str = "unidirectional"  # "unidirectional" | "bidirectional"
    enabled: bool = True
    category: str = ""
    # For unidirectional
    source_sheet: Optional[str] = None
    source_header: Optional[str] = None
    target_sheet: Optional[str] = None
    target_header: Optional[str] = None
    # For bidirectional
    sheet_a: Optional[str] = None
    header_a: Optional[str] = None
    sheet_b: Optional[str] = None
    header_b: Optional[str] = None
    # External sync
    is_external: bool = False
    target_doc_id: Optional[str] = None


class RuleUpdateRequest(BaseModel):
    """Request to update an existing sync rule."""
    enabled: Optional[bool] = None
    category: Optional[str] = None
    mode: Optional[str] = None
    source_sheet: Optional[str] = None
    source_header: Optional[str] = None
    target_sheet: Optional[str] = None
    target_header: Optional[str] = None
    sheet_a: Optional[str] = None
    header_a: Optional[str] = None
    sheet_b: Optional[str] = None
    header_b: Optional[str] = None
    is_external: Optional[bool] = None
    target_doc_id: Optional[str] = None


class RuleToggleRequest(BaseModel):
    """Request to toggle rule enabled status."""
    enabled: bool

# ============== Request/Response Models ==============

class SyncRowRequest(BaseModel):
    """Request to sync a single row."""

    project: str
    article: str
    source_sheet: str
    target_sheets: Optional[List[str]] = None
    spreadsheet_id: Optional[str] = None # Added support for direct ID


class SyncRangeRequest(BaseModel):
    """Request to sync a range."""

    project: str
    source_sheet: str
    range: str

class AddArticleRequest(BaseModel):
    project: str
    article: str
    spreadsheet_id: Optional[str] = None

class DeleteArticlesRequest(BaseModel):
    project: str
    articles: List[str]
    spreadsheet_id: Optional[str] = None

class SyncEventRequest(BaseModel):
    """Request from GAS onEdit trigger."""
    spreadsheet_id: str
    sheet_name: str
    row: int
    col: int
    value: Optional[Any] = None
    old_value: Optional[Any] = None
    user_email: Optional[str] = None
    header_name: Optional[str] = None
    row_key: Optional[str] = None
    # Защита от циклов
    sync_origin: str = "user"  # "user" | "sync"
    transaction_id: Optional[str] = None

class LogInitRequest(BaseModel):
    """Request to initialize logs."""
    spreadsheet_id: str

class SyncBatchEventRequest(BaseModel):
    """Request for batch onEdit events."""
    spreadsheet_id: str
    events: List[SyncEventRequest]


class PriceProcessRequest(BaseModel):
    """Request to process price file."""

    spreadsheet_id: str
    mode: str = "main"  # main, tester, samples, probes
    source_doc_id: Optional[str] = None
    dry_run: bool = False


class SortRequest(BaseModel):
    """Request to sort a sheet."""
    project: str = "Common" # specific project identifier if needed
    spreadsheet_id: str
    sheet_name: str
    column_name: str
    ascending: bool = True


class AIAnalyzeRequest(BaseModel):
    """Request for AI analysis."""
    spreadsheet_id: str
    sheet_name: str = "Информация"
    row_number: Optional[int] = None
    pdf_url: Optional[str] = None
    purpose: Optional[str] = None
    application: Optional[str] = None


class AIAnalyzeBatchRequest(BaseModel):
    """Request for batch AI analysis."""
    spreadsheet_id: str
    sheet_name: str = "Информация"
    delay_between: float = 2.0


class AICheckServiceRequest(BaseModel):
    """Request to check AI service status."""
    pass


class AIPdfAnalyzeRequest(BaseModel):
    """Request to analyze a PDF directly."""
    pdf_url: str
    purpose: Optional[str] = None
    application: Optional[str] = None


class AISimpleAnalyzeRequest(BaseModel):
    """Request for simple AI analysis (no PDF needed)."""
    product_name: str
    inci_text: Optional[str] = None
    purpose: Optional[str] = None
    application: Optional[str] = None


class TaskStatusResponse(BaseModel):
    """Task status response."""

    task_id: str
    status: str  # pending, running, completed, failed
    current_phase: Optional[int] = None
    total_phases: Optional[int] = None
    progress_percent: Optional[int] = None
    result: Optional[dict] = None
    error: Optional[str] = None


# ============== Sync Endpoints ==============

@api_router.post("/sync/row")
async def sync_row(request: SyncRowRequest):
    """Sync a single row by article."""
    logger.info(
        "sync_row_requested",
        project=request.project,
        article=request.article,
    )

    spreadsheet_id = request.spreadsheet_id
    if not spreadsheet_id:
        # Resolve ID from project name
        # Invert PROJECT_IDS
        project_map = {v: k for k, v in PROJECT_IDS.items()}
        spreadsheet_id = project_map.get(request.project)
        
    if not spreadsheet_id:
        raise HTTPException(status_code=400, detail=f"Could not resolve spreadsheet_id for project {request.project}")

    try:
        logging_service.add_log(spreadsheet_id, "API", "Запрос синхронизации строки", f"Артикул: {request.article}, Проект: {request.project}", "🚀 START")
        result = sync_service.sync_row(spreadsheet_id, request.project, request.article, request.source_sheet)
        logging_service.add_log(spreadsheet_id, "API", "Синхронизация строки завершена", f"Артикул: {request.article}", "✅ OK")
        return {
            "task_id": f"row_{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "status": "success",
            "message": f"Row sync processed for {request.article}",
            "details": result
        }
    except Exception as e:
        logger.error("sync_row_endpoint_failed", error=str(e))
        logging_service.add_log(spreadsheet_id, "API", "Ошибка синхронизации строки", str(e), "❌ ERR")
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/sync/range")
async def sync_range(request: SyncRangeRequest):
    """Sync a range of cells."""
    # For now, just alias to logging as range sync is complex to parse without context
    logger.info(
        "sync_range_requested",
        project=request.project,
        sheet=request.source_sheet,
        range=request.range,
    )

    return {
        "task_id": f"range_{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "status": "not_implemented", 
        "message": "Range sync not yet supported via Python"
    }


@api_router.post("/sync/full")
async def sync_full(project: str, source_sheet: str, spreadsheet_id: Optional[str] = None):
    """Full sync of a sheet."""
    logger.info("full_sync_requested", project=project, sheet=source_sheet)
    
    if not spreadsheet_id:
        project_map = {v: k for k, v in PROJECT_IDS.items()}
        spreadsheet_id = project_map.get(project)
        
    if not spreadsheet_id:
        raise HTTPException(status_code=400, detail=f"Could not resolve spreadsheet_id for project {project}")

    try:
        result = sync_service.sync_full(spreadsheet_id, project, source_sheet)
        return {
            "task_id": f"full_{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "status": "success", 
            "details": result
        }
    except Exception as e:
        logger.error("sync_full_endpoint_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/sync/add-article")
async def add_article(request: AddArticleRequest):
    """Add new article to all relevant sheets."""
    logger.info("add_article_requested", project=request.project, article=request.article)
    
    spreadsheet_id = request.spreadsheet_id
    project_key = request.project
    
    # Resolve spreadsheet_id from project if missing
    if not spreadsheet_id and project_key != "UNKNOWN":
        project_map = {v: k for k, v in PROJECT_IDS.items()}
        spreadsheet_id = project_map.get(project_key)
        
    # Resolve project_key from spreadsheet_id if UNKNOWN
    if project_key == "UNKNOWN" and spreadsheet_id:
        project_key = PROJECT_IDS.get(spreadsheet_id, "UNKNOWN")

    if not spreadsheet_id:
        raise HTTPException(status_code=400, detail=f"Could not resolve spreadsheet_id")

    try:
        logging_service.add_log(spreadsheet_id, "API", "Запрос на добавление артикула", f"Артикул: {request.article}", "🚀 START")
        details = await sync_service.add_article(spreadsheet_id, request.article, project_key)
        return {
            "status": "success",
            "message": "Артикул обработан",
            "details": details
        }
    except Exception as e:
        logger.error("add_article_endpoint_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/sync/delete-articles")
async def delete_articles(request: DeleteArticlesRequest):
    """Delete articles from relevant sheets."""
    logger.info("delete_articles_requested", project=request.project, count=len(request.articles))
    
    spreadsheet_id = request.spreadsheet_id
    if not spreadsheet_id:
        project_map = {v: k for k, v in PROJECT_IDS.items()}
        spreadsheet_id = project_map.get(request.project)
        
    if not spreadsheet_id:
        raise HTTPException(status_code=400, detail=f"Could not resolve spreadsheet_id for project {request.project}")

    try:
        logging_service.add_log(spreadsheet_id, "API", "Запрос на удаление артикулов", f"Количество: {len(request.articles)}", "🚀 START")
        result = sync_service.delete_articles(spreadsheet_id, request.articles)
        return {
            "status": "success",
            "message": f"Deleted {result.get('total_deleted')} rows total",
            "details": result
        }
    except Exception as e:
        logger.error("delete_articles_endpoint_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/sync/event")
async def sync_event(request: SyncEventRequest, background_tasks: BackgroundTasks):
    """
    Handle onEdit sync event (async).
    Uses background tasks for non-blocking execution.
    """
    logger.info(
        "sync_event_received",
        sheet=request.sheet_name,
        row=request.row,
        col=request.col,
        header=request.header_name
    )
    
    try:
        # Construct event data dict
        event_data = {
            "sheet_name": request.sheet_name,
            "row": request.row,
            "col": request.col,
            "value": request.value,
            "old_value": request.old_value,
            "user_email": request.user_email,
            "header_name": request.header_name,
            "row_key": request.row_key,
            "sync_origin": request.sync_origin,
            "transaction_id": request.transaction_id,
        }
        
        # Queue task
        background_tasks.add_task(sync_service.sync_event, request.spreadsheet_id, event_data)
        
        return {
            "status": "queued",
            "message": "Sync queued"
        }
    except Exception as e:
        logger.error("sync_event_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/sync/batch-event")
async def sync_batch_event(request: SyncBatchEventRequest, background_tasks: BackgroundTasks):
    """
    Handle batch onEdit sync events (async).
    """
    logger.info("batch_sync_event_received", count=len(request.events))
    
    try:
        count = 0
        for event in request.events:
            # Construct event data dict
            event_data = {
                "sheet_name": event.sheet_name,
                "row": event.row,
                "col": event.col,
                "value": event.value,
                "old_value": event.old_value,
                "user_email": event.user_email,
                "header_name": event.header_name,
                "row_key": event.row_key,
                "sync_origin": event.sync_origin,
                "transaction_id": event.transaction_id,
            }
            
            background_tasks.add_task(sync_service.sync_event, request.spreadsheet_id, event_data)
            count += 1
        
        logging_service.add_log(
            request.spreadsheet_id, 
            "СИНХРОНИЗАЦИЯ", 
            "Получена пачка событий (batch)", 
            f"Количество: {count}", 
            "📥 В ОЧЕРЕДИ"
        )
        
        return {
            "status": "queued",
            "message": f"Queued {count} events"
        }
    except Exception as e:
        logger.error("sync_batch_event_failed", error=str(e))
        logging_service.add_log(
            request.spreadsheet_id,
            "СИНХРОНИЗАЦИЯ",
            f"Ошибка при обработке пачки событий (batch)",
            f"Количество: {len(request.events)}, Ошибка: {str(e)}",
            "❌ ОШИБКА"
        )
        raise HTTPException(status_code=500, detail=str(e))


# ... (SortRequest model remains)
# ... (SortRequest model remains)

# ============== Rules Management ==============

@api_router.get("/rules/{spreadsheet_id}")
async def get_rules(spreadsheet_id: str, force_reload: bool = False, include_disabled: bool = True):
    """Return sync rules (server-side YAML storage)."""
    try:
        rules = sync_service.list_rules(
            spreadsheet_id,
            force_reload=force_reload,
            include_disabled=include_disabled
        )
        return {"rules": rules}
    except Exception as e:
        logger.error("rules_get_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/rules/{spreadsheet_id}/reload")
async def reload_rules(spreadsheet_id: str, include_disabled: bool = True):
    """Force reload rules from sheets/storage and refresh cache."""
    try:
        rules = sync_service.list_rules(
            spreadsheet_id,
            force_reload=True,
            include_disabled=include_disabled
        )
        return {"status": "ok", "rules": rules, "message": "Rules reloaded and cache refreshed"}
    except Exception as e:
        logger.error("rules_reload_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/rules/{spreadsheet_id}/create")
async def create_rule_endpoint(spreadsheet_id: str, request: RuleCreateRequest):
    """Create a new sync rule."""
    try:
        rule = sync_service.create_rule(spreadsheet_id, request.model_dump())
        return {"status": "ok", "rule": rule}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error("rules_create_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.patch("/rules/{spreadsheet_id}/{rule_id}")
async def update_rule_endpoint(spreadsheet_id: str, rule_id: str, request: RuleUpdateRequest):
    """Update an existing sync rule."""
    try:
        payload = request.model_dump(exclude_none=True)
        rule = sync_service.update_rule(spreadsheet_id, rule_id, payload)
        return {"status": "ok", "rule": rule}
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        logger.error("rules_update_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.delete("/rules/{spreadsheet_id}/{rule_id}")
async def delete_rule_endpoint(spreadsheet_id: str, rule_id: str):
    """Delete a sync rule by ID."""
    try:
        sync_service.delete_rule(spreadsheet_id, rule_id)
        return {"status": "ok", "deleted": True}
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        logger.error("rules_delete_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.patch("/rules/{spreadsheet_id}/{rule_id}/toggle")
async def toggle_rule_endpoint(spreadsheet_id: str, rule_id: str, request: RuleToggleRequest):
    """Toggle rule enabled/disabled."""
    try:
        rule = sync_service.toggle_rule(spreadsheet_id, rule_id, request.enabled)
        return {"status": "ok", "rule": rule}
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        logger.error("rules_toggle_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.get("/rules-ui")
async def get_rules_ui():
    """Serve the Rule Manager UI."""
    return FileResponse("config/rule_manager.html")


@api_router.get("/logs-ui")
async def get_logs_ui():
    """Serve the Logs Dashboard UI."""
    return FileResponse("config/logs_manager.html")



# ============== Sync Logs ==============

def _parse_iso_datetime(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        raise HTTPException(status_code=400, detail=f"Invalid datetime format: {value}")


@api_router.get("/sync-logs/{spreadsheet_id}")
async def list_sync_logs(
    spreadsheet_id: str,
    limit: int = 200,
    offset: int = 0,
    order: str = "desc",
    project: Optional[str] = None,
    category: Optional[str] = None,
    status: Optional[str] = None,
    row_key: Optional[str] = None,
    source: Optional[str] = None,
    target: Optional[str] = None,
    rule_id: Optional[str] = None,
    start: Optional[str] = None,
    end: Optional[str] = None,
):
    """List sync journal entries from server storage."""
    try:
        start_dt = _parse_iso_datetime(start)
        end_dt = _parse_iso_datetime(end)
        safe_limit = min(max(limit, 0), 1000)
        safe_offset = max(offset, 0)

        items, total = sync_log_service.list_entries(
            spreadsheet_id=spreadsheet_id,
            limit=safe_limit,
            offset=safe_offset,
            order=order,
            project=project,
            category=category,
            status=status,
            row_key=row_key,
            source=source,
            target=target,
            rule_id=rule_id,
            start=start_dt,
            end=end_dt,
        )

        return {
            "items": items,
            "total": total,
            "limit": safe_limit,
            "offset": safe_offset,
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error("sync_logs_list_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.get("/sync-logs/{spreadsheet_id}/stats")
async def sync_logs_stats(spreadsheet_id: str):
    """Get statistics for sync logs (summary by category and status)."""
    try:
        items, _ = sync_log_service.list_entries(
            spreadsheet_id=spreadsheet_id,
            limit=10000,  # Get all for statistics
        )

        # Aggregate statistics
        stats = {
            "total": len(items),
            "by_category": {},
            "by_status": {},
            "by_rule_id": {},
            "latest_timestamp": None,
            "oldest_timestamp": None,
        }

        for entry in items:
            # By category
            cat = entry.get("category", "unknown")
            if cat not in stats["by_category"]:
                stats["by_category"][cat] = 0
            stats["by_category"][cat] += 1

            # By status
            status = entry.get("status", "unknown")
            if status not in stats["by_status"]:
                stats["by_status"][status] = 0
            stats["by_status"][status] += 1

            # By rule_id
            rule_id = entry.get("rule_id", "unknown")
            if rule_id not in stats["by_rule_id"]:
                stats["by_rule_id"][rule_id] = 0
            stats["by_rule_id"][rule_id] += 1

            # Timestamps
            ts_str = entry.get("timestamp")
            if ts_str:
                if not stats["latest_timestamp"]:
                    stats["latest_timestamp"] = ts_str
                if not stats["oldest_timestamp"]:
                    stats["oldest_timestamp"] = ts_str

        return stats
    except Exception as e:
        logger.error("sync_logs_stats_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.get("/sync-logs/{spreadsheet_id}/export")
async def export_sync_logs(
    spreadsheet_id: str,
    format: str = "json",
    limit: int = 10000,
):
    """Export sync logs in JSON or CSV format."""
    try:
        items, _ = sync_log_service.list_entries(
            spreadsheet_id=spreadsheet_id,
            limit=min(limit, 50000),
        )

        if format.lower() == "csv":
            import csv
            import io
            import tempfile
            from pathlib import Path

            output = io.StringIO()
            if items:
                fieldnames = list(items[0].keys())
                writer = csv.DictWriter(output, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(items)

            # Write to temp file for FileResponse
            with tempfile.NamedTemporaryFile(
                mode='w', suffix='.csv', delete=False, encoding='utf-8'
            ) as f:
                f.write(output.getvalue())
                temp_path = f.name

            return FileResponse(
                path=temp_path,
                filename=f"sync_logs_{spreadsheet_id}.csv",
                media_type="text/csv",
            )
        else:  # JSON
            return {
                "format": "json",
                "spreadsheet_id": spreadsheet_id,
                "count": len(items),
                "items": items,
            }
    except Exception as e:
        logger.error("sync_logs_export_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/sync-logs/{spreadsheet_id}/truncate")
async def truncate_sync_logs(
    spreadsheet_id: str,
    keep_days: int = 0,
):
    """Remove old sync logs (older than keep_days)."""
    try:
        items, _ = sync_log_service.list_entries(
            spreadsheet_id=spreadsheet_id,
            limit=100000,
        )

        if keep_days <= 0:
            # Delete all
            deleted = len(items)
            log_path = sync_log_service._log_path(spreadsheet_id)
            lock_path = sync_log_service._lock_path(spreadsheet_id)
            from filelock import FileLock
            with FileLock(lock_path):
                if log_path.exists():
                    log_path.unlink()
        else:
            # Keep entries newer than keep_days
            from datetime import datetime, timedelta
            cutoff = datetime.utcnow() - timedelta(days=keep_days)
            filtered = []
            deleted = 0

            for entry in items:
                ts = sync_log_service._parse_timestamp(entry.get("timestamp"))
                if ts and ts >= cutoff:
                    filtered.append(entry)
                else:
                    deleted += 1

            # Write back filtered logs
            log_path = sync_log_service._log_path(spreadsheet_id)
            lock_path = sync_log_service._lock_path(spreadsheet_id)
            from filelock import FileLock
            import json

            with FileLock(lock_path):
                with open(log_path, "w", encoding="utf-8") as handle:
                    for entry in filtered:
                        handle.write(json.dumps(entry, ensure_ascii=False) + "\n")

        return {
            "deleted": deleted,
            "remaining": len(items) - deleted,
        }
    except Exception as e:
        logger.error("sync_logs_truncate_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))




# --- Log Management Endpoints ---

@api_router.post("/logs/recreate")
async def recreate_log_sheet_endpoint(request: LogInitRequest):
    """Recreate log sheet (ensure headers are correct)."""
    try:
        logging_service.recreate_log_sheet(request.spreadsheet_id, force_clear=False)
        return {"status": "success", "message": "Log sheet recreated"}
    except Exception as e:
        logger.error("log_recreate_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@api_router.post("/logs/clean")
async def clean_log_sheet_endpoint(request: LogInitRequest):
    """Clean log sheet (force clear and recreate headers)."""
    try:
        logging_service.recreate_log_sheet(request.spreadsheet_id, force_clear=True)
        return {"status": "success", "message": "Log sheet cleaned"}
    except Exception as e:
        logger.error("log_clean_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@api_router.post("/logs/recreate-debug")
async def recreate_debug_log_sheet_endpoint(request: LogInitRequest):
    """Recreate debug log sheet ('Журнал логов')."""
    try:
        # 'Журнал логов' is the sheet name for debug logs
        logging_service.recreate_log_sheet(request.spreadsheet_id, force_clear=False, sheet_name="Журнал логов")
        return {"status": "success", "message": "Debug log sheet recreated"}
    except Exception as e:
        logger.error("log_debug_recreate_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@api_router.post("/logs/archive")
async def archive_logs_endpoint(request: LogInitRequest):
    """Manually archive logs."""
    try:
        # Need to determine project prefix, default to 'Common' or try to get from request/sheet?
        # For now, let's look up project from sheet title or just use "Project"
        # Ideally, we should receive 'project' in request, but LogInitRequest only has spreadsheet_id
        # We can fetch project from sheet name or config.
        # Let's peek at sheet name? Or just use "Archive"
        
        # Simple heuristic:
        try:
             ss = sheets_service.get_spreadsheet(request.spreadsheet_id)
             title = ss.title.lower()
             prefix = "common"
             if "mt" in title: prefix = "mt"
             elif "sk" in title: prefix = "sk"
             elif "ss" in title: prefix = "ss"
        except:
             prefix = "common"
             
        result = await logging_service.archive_logs(request.spreadsheet_id, project_prefix=prefix)
        return {"status": "success", "data": result}
    except Exception as e:
        logger.error("log_archive_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@api_router.get("/logs/archive/status")
async def get_archive_status_endpoint(spreadsheet_id: str):
    """Get archive status."""
    try:
        try:
             ss = sheets_service.get_spreadsheet(spreadsheet_id)
             title = ss.title.lower()
             prefix = "common"
             if "mt" in title: prefix = "mt"
             elif "sk" in title: prefix = "sk"
             elif "ss" in title: prefix = "ss"
        except:
             prefix = "common"
             
        status = logging_service.get_archive_status(project_prefix=prefix)
        return {"status": "success", "data": status}
    except Exception as e:
        logger.error("get_archive_status_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

# --- Document Collection Endpoints ---

# --- Document Collection Endpoints ---

class CollectDocumentsRequest(BaseModel):
    spreadsheet_id: str
    target_sheet: str = "Для инвойса"
    brand_prefix: str = "MT"

@api_router.post("/documents/structure-353pp")
async def structure_documents_353pp_endpoint(request: CollectDocumentsRequest):
    """Structure documents for 353pp application (from New Sert sheet)."""
    try:
        result = await certification_service.structure_documents_353pp(
            request.spreadsheet_id
        )
        return {"status": "success", "data": result}
    except Exception as e:
        logger.error("structure_353pp_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@api_router.post("/documents/collect")
async def collect_documents_endpoint(request: CollectDocumentsRequest):
    """Collect and copy documents for invoice."""
    try:
        # Get invoice service
        inv_service = get_invoice_service(sheets_service)
        
        result = await inv_service.collect_and_copy_documents(
            spreadsheet_id=request.spreadsheet_id,
            target_sheet=request.target_sheet,
            brand_prefix=request.brand_prefix
        )
        return {"status": "success", "data": result}
    except Exception as e:
        logger.error("collect_documents_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

# --- Function Logs Endpoints ---

@api_router.get("/function-logs/executions")
async def list_function_executions(
    module: Optional[str] = None,
    function_name: Optional[str] = None,
    limit: int = 100,
    offset: int = 0,
):
    """List historical function executions."""
    try:
        executions, total = function_log_service.list_executions(
            module=module,
            function_name=function_name,
            limit=limit,
            offset=offset,
        )
        return {
            "executions": executions,
            "total": total,
            "limit": limit,
            "offset": offset,
        }
    except Exception as e:
        logger.error("function_logs_list_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# --- Metadata & External Docs Endpoints ---

class ExternalDocAddRequest(BaseModel):
    name: str
    doc_id: str

@api_router.get("/meta/{spreadsheet_id}/sheets")
async def get_meta_sheets(spreadsheet_id: str):
    """List sheets in a spreadsheet."""
    try:
        sh = sheets_service.gc.open_by_key(spreadsheet_id)
        return {"sheets": [s.title for s in sh.worksheets()]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@api_router.get("/meta/{spreadsheet_id}/{sheet_name}/headers")
async def get_meta_headers(spreadsheet_id: str, sheet_name: str):
    """List column headers in a sheet."""
    try:
        headers = sheets_service.get_worksheet_headers(spreadsheet_id, sheet_name)
        return {"headers": headers}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@api_router.get("/meta/discovery")
async def get_meta_discovery():
    """Build and return a map of common headers across the ecosystem."""
    return meta_cache_service.get_discovery_map()

@api_router.post("/meta/index")
async def trigger_meta_index(background_tasks: BackgroundTasks, spreadsheet_id: Optional[str] = None):
    """
    Trigger re-indexing of metadata. 
    If spreadsheet_id is provided, only index that document.
    Otherwise, index all projects in PROJECT_IDS.
    """
    if spreadsheet_id:
        targets = [spreadsheet_id]
    else:
        # Get all unique IDs from PROJECT_IDS
        targets = list(PROJECT_IDS.keys())
    
    background_tasks.add_task(meta_cache_service.build_global_index, targets)
    return {"status": "queued", "targets": len(targets)}

@api_router.get("/meta/index/status")
async def get_meta_index_status():
    """Get status of the metadata cache."""
    return {
        "updated_at": meta_cache_service.index.get("updated_at"),
        "total_spreadsheets": len(meta_cache_service.index.get("spreadsheets", {})),
        "total_headers": len(meta_cache_service.index.get("headers", {})),
        "spreadsheets": meta_cache_service.get_indexed_spreadsheets()
    }

@api_router.get("/meta/search")
async def search_meta_header(q: str):
    """Search for sheets containing a specific header."""
    return meta_cache_service.search_header(q)

@api_router.get("/external-docs")
async def list_external_docs():
    """List registered external documents."""
    return {"docs": external_docs_service.list_docs()}

@api_router.post("/external-docs")
async def add_external_doc(request: ExternalDocAddRequest):
    """Register a new external document."""
    try:
        doc = external_docs_service.add_doc(request.name, request.doc_id)
        return {"status": "ok", "doc": doc}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@api_router.delete("/external-docs/{doc_id}")
async def remove_external_doc(doc_id: str):
    """Unregister an external document."""
    if external_docs_service.remove_doc(doc_id):
        return {"status": "ok"}
    raise HTTPException(status_code=404, detail="Document not found")


@api_router.post("/rules/{spreadsheet_id}")
async def save_rules(spreadsheet_id: str, request: RulesSaveRequest):
    """
    Replace rules for a spreadsheet. IDs пересоздаются автоматически в формате 001-<SRC>-<TGT>(<Header>).
    """
    try:
        payload = [r.model_dump() for r in request.rules]
        rules = sync_service.save_rules(spreadsheet_id, payload, validate_headers=True)
        return {"status": "ok", "rules": rules}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error("rules_save_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

class LoadFunctionsRequest(BaseModel):
    """Request to run load functions."""
    spreadsheet_id: str
    project: str = "Common"

@api_router.post("/sort")
async def sort_sheet(request: SortRequest):
    """Sort a sheet by column."""
    logger.info(
        "sort_requested",
        spreadsheet_id=request.spreadsheet_id,
        sheet=request.sheet_name,
        column=request.column_name
    )
    
    try:
        logging_service.add_log(
            request.spreadsheet_id,
            "СОРТИРОВКА",
            f"Старт сортировки листа {request.sheet_name}",
            f"Колонка: {request.column_name}, По возрастанию: {request.ascending}",
            "🔄"
        )
        sheets_service.sort_by_header(
            request.spreadsheet_id, 
            request.sheet_name, 
            request.column_name, 
            request.ascending
        )
        logging_service.add_log(
            request.spreadsheet_id,
            "СОРТИРОВКА",
            f"Сортировка листа {request.sheet_name}",
            f"Колонка: {request.column_name}, По возрастанию: {request.ascending}",
            "✅ УСПЕХ"
        )
        return {"status": "success", "message": f"Sorted by {request.column_name}"}
    except Exception as e:
        logger.error("sort_endpoint_error", error=str(e))
        logging_service.add_log(
            request.spreadsheet_id,
            "СОРТИРОВКА",
            f"Ошибка сортировки листа {request.sheet_name}",
            str(e),
            "❌"
        )
        raise HTTPException(status_code=500, detail=str(e))


class StructureSortRequest(BaseModel):
    """Request for high-performance structure sorting (replaces GAS structureMultipleSheets)."""
    spreadsheet_id: str
    mode: str  # 'byManufacturer' or 'byPrice'


@api_router.post("/sort/structure")
async def sort_structure(request: StructureSortRequest):
    """
    High-performance structure sorting.
    Sorts multiple sheets (Заказ, Динамика цены, Расчет цены) by grouping rows.
    Replaces the slow GAS structureMultipleSheets function.
    """
    logger.info(
        "structure_sort_requested",
        spreadsheet_id=request.spreadsheet_id,
        mode=request.mode
    )
    
    if request.mode not in ["byManufacturer", "byPrice"]:
        raise HTTPException(status_code=400, detail="Invalid mode. Use 'byManufacturer' or 'byPrice'")
    
    try:
        logging_service.add_log(
            request.spreadsheet_id,
            "СТРУКТУРА",
            f"Старт структурной сортировки ({request.mode})",
            "Листы: Заказ, Динамика цены, Расчет цены",
            "🔄"
        )
        result = sorting_service.sort_sheets(request.spreadsheet_id, request.mode)
        return {
            "status": "success",
            "message": f"Sorted {len(result['sheets_processed'])} sheets",
            "details": result
        }
    except Exception as e:
        logger.error("structure_sort_failed", error=str(e))
        logging_service.add_log(
            request.spreadsheet_id,
            "СТРУКТУРА",
            f"Ошибка структурной сортировки ({request.mode})",
            str(e),
            "❌"
        )
        raise HTTPException(status_code=500, detail=str(e))

@api_router.post("/load-functions")
async def run_load_functions(request: LoadFunctionsRequest):
    """Run heavy logic normally executed on sheet load."""
    logger.info("load_functions_requested", spreadsheet_id=request.spreadsheet_id)
    
    try:
        result = sheets_service.process_load_logic(request.spreadsheet_id)
        return {
            "status": "success", 
            "message": "Values and formulas updated successfully",
            "details": result
        }
    except Exception as e:
        logger.error("load_functions_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== Price Processing ==============

@api_router.post("/price/process/{project}")
async def process_price(
    project: str,
    request: PriceProcessRequest,
    background_tasks: BackgroundTasks
):
    """
    Process supplier price list (Б/З поставщик).

    Args:
        project: Project code (mt, sk, ss)
        request: Processing parameters (spreadsheet_id, mode, dry_run)

    Returns:
        Processing result or preview if dry_run=True
    """
    logger.info(
        "price_process_requested",
        project=project,
        mode=request.mode,
        spreadsheet_id=request.spreadsheet_id,
        dry_run=request.dry_run
    )

    try:
        # For dry_run, process synchronously to return preview
        if request.dry_run:
            result = await price_processor.process(
                project=project,
                mode=request.mode,
                spreadsheet_id=request.spreadsheet_id,
                source_doc_id=request.source_doc_id,
                dry_run=True
            )
            return result

        # For actual processing, run in background
        background_tasks.add_task(
            price_processor.process,
            project=project,
            mode=request.mode,
            spreadsheet_id=request.spreadsheet_id,
            source_doc_id=request.source_doc_id,
            dry_run=False
        )

        return {
            "status": "queued",
            "message": f"Processing {project.upper()} {request.mode} started",
            "task_id": f"price_{project}_{request.mode}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        }

    except Exception as e:
        logger.error("price_process_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.get("/price/status/{task_id}", response_model=TaskStatusResponse)
async def get_price_status(task_id: str):
    """Get price processing status."""
    # TODO: Get actual status from task queue
    return TaskStatusResponse(
        task_id=task_id,
        status="running",
        current_phase=3,
        total_phases=9,
        progress_percent=33,
    )


@api_router.post("/price/cancel/{task_id}")
async def cancel_price_processing(task_id: str):
    """Cancel price processing."""
    logger.info("price_cancel_requested", task_id=task_id)

    # TODO: Cancel task
    return {"task_id": task_id, "status": "cancelled"}


# ============== Cascade Processing ==============

class CascadeProcessRequest(BaseModel):
    """Request for cascade processing."""
    spreadsheet_id: str
    sheet_name: str = "Сертификация"
    row: Optional[int] = None
    changed_column: Optional[str] = None
    new_value: Optional[str] = None
    dry_run: bool = False


class CascadeProcessResponse(BaseModel):
    """Response from cascade processing."""
    status: str
    row: int = 0
    changes: List[Dict[str, Any]] = []
    applied: bool = False
    message: str = ""


@api_router.post("/cascade/process", response_model=CascadeProcessResponse)
async def process_cascade(request: CascadeProcessRequest):
    """
    Process cascade rules for certification sheet.

    When trigger columns change (Наименования рус по ДС, Наименования англ по ДС,
    Объём, Код ТН ВЭД), recalculates derived fields:
    - Объём англ. (volume in English)
    - Наименование ДС (combined name)
    - Наименование для инвойса (Russian invoice name)
    - Наименование для инвойса Англ (English invoice name)

    Args:
        request: Cascade processing parameters

    Returns:
        List of changes made or preview if dry_run=True
    """
    from src.services.cascade_processor import get_cascade_processor

    logger.info(
        "cascade_process_requested",
        spreadsheet_id=request.spreadsheet_id,
        sheet_name=request.sheet_name,
        row=request.row,
        column=request.changed_column,
        dry_run=request.dry_run
    )

    try:
        processor = get_cascade_processor(sheets_service)

        if request.row and request.changed_column:
            # Process single row
            result = await processor.process_single_row(
                spreadsheet_id=request.spreadsheet_id,
                sheet_name=request.sheet_name,
                row=request.row,
                changed_column=request.changed_column,
                new_value=request.new_value,
                dry_run=request.dry_run
            )

            return CascadeProcessResponse(
                status="success" if result.applied else "no_changes",
                row=result.row,
                changes=[
                    {
                        "column": c.column,
                        "old_value": c.old_value,
                        "new_value": c.new_value
                    }
                    for c in result.changes
                ],
                applied=result.applied,
                message=result.message
            )
        else:
            # Process all rows
            summary = await processor.process_all_rows(
                spreadsheet_id=request.spreadsheet_id,
                sheet_name=request.sheet_name,
                dry_run=request.dry_run
            )

            return CascadeProcessResponse(
                status=summary.get("status", "error"),
                changes=[],
                applied=summary.get("changed", 0) > 0,
                message=f"Processed {summary.get('processed', 0)} rows, {summary.get('changed', 0)} changed"
            )

    except Exception as e:
        logger.error("cascade_process_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/cascade/recalculate-all")
async def recalculate_all_cascades(request: CascadeProcessRequest):
    """
    Recalculate cascades for all rows in certification sheet.
    Equivalent to GAS runManualCascadeOnCertification().
    """
    from src.services.cascade_processor import get_cascade_processor

    logger.info(
        "cascade_recalculate_all_requested",
        spreadsheet_id=request.spreadsheet_id,
        dry_run=request.dry_run
    )

    try:
        processor = get_cascade_processor(sheets_service)

        result = await processor.process_all_rows(
            spreadsheet_id=request.spreadsheet_id,
            sheet_name=request.sheet_name,
            dry_run=request.dry_run
        )

        return {
            "status": result.get("status", "error"),
            "total_rows": result.get("total_rows", 0),
            "processed": result.get("processed", 0),
            "changed": result.get("changed", 0),
            "dry_run": result.get("dry_run", False),
            "errors": result.get("errors", [])
        }

    except Exception as e:
        logger.error("cascade_recalculate_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== Order Stages ==============

class OrderFilterRequest(BaseModel):
    """Request for order stage filtering."""
    spreadsheet_id: str
    stage: str = "all"  # all, order, promotions, set, price
    sheet_name: str = "Заказ"
    dry_run: bool = False


class OrderFilterResponse(BaseModel):
    """Response from order stage filtering."""
    status: str
    stage: str
    visible_rows: int = 0
    hidden_rows: int = 0
    hidden_columns: int = 0
    message: str = ""


@api_router.post("/order/filter", response_model=OrderFilterResponse)
async def filter_order_stage(request: OrderFilterRequest):
    """
    Filter order sheet by stage.

    Stages:
    - all: Show all data (remove all filters)
    - order: Show order-related columns
    - promotions: Show promotion-related columns
    - set: Show set-related columns
    - price: Show price-related columns

    Each stage hides specific columns and filters rows by status.
    """
    from src.services.order_service import get_order_service, OrderStageType

    logger.info(
        "order_filter_requested",
        spreadsheet_id=request.spreadsheet_id,
        stage=request.stage,
        dry_run=request.dry_run
    )

    try:
        # Validate stage
        try:
            stage = OrderStageType(request.stage)
        except ValueError:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid stage: {request.stage}. Valid values: all, order, promotions, set, price"
            )

        service = get_order_service(sheets_service)

        result = await service.apply_stage_filter(
            spreadsheet_id=request.spreadsheet_id,
            stage=stage,
            sheet_name=request.sheet_name,
            dry_run=request.dry_run
        )

        return OrderFilterResponse(
            status="success",
            stage=result.stage,
            visible_rows=result.visible_rows,
            hidden_rows=result.hidden_rows,
            hidden_columns=result.hidden_columns,
            message=result.message
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("order_filter_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/order/show-all")
async def show_all_order_data(request: OrderFilterRequest):
    """
    Show all data on order sheet (remove all filters).
    Equivalent to GAS showAllOrderData().
    """
    request.stage = "all"
    return await filter_order_stage(request)


# ============== Data Export ==============

class ExportRequest(BaseModel):
    """Request for data export."""
    spreadsheet_id: str
    project: str  # mt, sk, ss
    target_doc_id: Optional[str] = None  # Override default target
    dry_run: bool = False


class ExportResponse(BaseModel):
    """Response from data export."""
    status: str
    export_type: str
    exported_rows: int = 0
    target_doc_id: str = ""
    target_sheet_name: str = ""
    target_url: str = ""
    message: str = ""


@api_router.post("/export/promotions", response_model=ExportResponse)
async def export_promotions(request: ExportRequest):
    """
    Export promotions data to target document.

    Reads rows from "Заказ" sheet where "АКЦИИ" column has value,
    enriches with data from "Главная" and "Сертификация",
    then writes to target document.

    Target documents by project:
    - SK: 1YkGP-1Ipn7qLMKJyxLtm3ATrOhCxO2OuFLMt5WK8tsg
    - SS: 1Q20jk9Cy8gIEJyKQ2-Ph34qqX3Y_oEdKAOK-o_oaFHQ
    - MT: 140vuIAJ1dcuAoc10T5EnIFjx1lUq7e7oroBJlBs1TDA
    """
    from src.services.export_service import get_export_service, ExportType

    logger.info(
        "export_promotions_requested",
        spreadsheet_id=request.spreadsheet_id,
        project=request.project,
        dry_run=request.dry_run
    )

    try:
        service = get_export_service(sheets_service)

        result = await service.export(
            spreadsheet_id=request.spreadsheet_id,
            project=request.project,
            export_type=ExportType.PROMOTIONS,
            target_doc_id=request.target_doc_id,
            dry_run=request.dry_run
        )

        return ExportResponse(
            status="success",
            export_type=result.export_type,
            exported_rows=result.exported_rows,
            target_doc_id=result.target_doc_id,
            target_sheet_name=result.target_sheet_name,
            target_url=result.target_url,
            message=result.message
        )

    except Exception as e:
        logger.error("export_promotions_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/export/sets", response_model=ExportResponse)
async def export_sets(request: ExportRequest):
    """
    Export sets data to target document.

    Reads rows from "Заказ" sheet where "Набор" column has value,
    enriches with data from "Главная" and "Сертификация",
    then writes to target document.
    """
    from src.services.export_service import get_export_service, ExportType

    logger.info(
        "export_sets_requested",
        spreadsheet_id=request.spreadsheet_id,
        project=request.project,
        dry_run=request.dry_run
    )

    try:
        service = get_export_service(sheets_service)

        result = await service.export(
            spreadsheet_id=request.spreadsheet_id,
            project=request.project,
            export_type=ExportType.SETS,
            target_doc_id=request.target_doc_id,
            dry_run=request.dry_run
        )

        return ExportResponse(
            status="success",
            export_type=result.export_type,
            exported_rows=result.exported_rows,
            target_doc_id=result.target_doc_id,
            target_sheet_name=result.target_sheet_name,
            target_url=result.target_url,
            message=result.message
        )

    except Exception as e:
        logger.error("export_sets_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== Certification ==============

class CertificationNewsRequest(BaseModel):
    """Request for creating news sheet."""
    spreadsheet_id: str
    source_sheet: str = "Сертификация"
    target_sheet: str = "New sert"
    dry_run: bool = False


class CertificationSpiritsRequest(BaseModel):
    """Request for spirit calculations."""
    spreadsheet_id: str
    sheet_name: str = "Сертификация"
    dry_run: bool = False


class CertificationProtocolsRequest(BaseModel):
    """Request for protocol generation."""
    spreadsheet_id: str
    protocol_type: str = "353pp"
    dry_run: bool = False


class CertificationResponse(BaseModel):
    """Response from certification operations."""
    status: str
    rows_affected: int = 0
    sheet_name: str = ""
    message: str = ""


@api_router.post("/certification/news-sheet", response_model=CertificationResponse)
async def create_news_sheet(request: CertificationNewsRequest):
    """
    Create "New sert" sheet from Certification sheet.

    Filters rows where status indicates new products pending certification.
    Creates a new sheet with summary of products needing attention.

    Equivalent to GAS createNewsSheetFromCertification().
    """
    from src.services.certification_service import get_certification_service

    logger.info(
        "certification_news_sheet_requested",
        spreadsheet_id=request.spreadsheet_id,
        dry_run=request.dry_run
    )

    try:
        service = get_certification_service(sheets_service)

        result = await service.create_news_sheet(
            spreadsheet_id=request.spreadsheet_id,
            source_sheet=request.source_sheet,
            target_sheet=request.target_sheet,
            dry_run=request.dry_run
        )

        return CertificationResponse(
            status=result.get("status", "error"),
            rows_affected=result.get("rows_created", 0),
            sheet_name=result.get("sheet_name", request.target_sheet),
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("certification_news_sheet_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/certification/spirits/calculate", response_model=CertificationResponse)
async def calculate_spirits(request: CertificationSpiritsRequest):
    """
    Calculate and assign spirit numbers to certification rows.

    Spirit numbers are required for products containing alcohol.

    Equivalent to GAS calculateAndAssignSpiritNumbers().
    """
    from src.services.certification_service import get_certification_service

    logger.info(
        "certification_spirits_calculate_requested",
        spreadsheet_id=request.spreadsheet_id,
        dry_run=request.dry_run
    )

    try:
        service = get_certification_service(sheets_service)

        result = await service.calculate_spirit_numbers(
            spreadsheet_id=request.spreadsheet_id,
            sheet_name=request.sheet_name,
            dry_run=request.dry_run
        )

        return CertificationResponse(
            status=result.get("status", "error"),
            rows_affected=result.get("calculated", 0),
            sheet_name=request.sheet_name,
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("certification_spirits_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/certification/protocols-353pp", response_model=CertificationResponse)
async def generate_protocols_353pp(request: CertificationProtocolsRequest):
    """
    Generate certification protocols for 353пп.

    NOTE: This endpoint is a placeholder. Full implementation requires
    Google Drive/Docs integration for document generation.

    Equivalent to GAS generateProtocols_353pp().
    """
    from src.services.certification_service import get_certification_service

    logger.info(
        "certification_protocols_requested",
        spreadsheet_id=request.spreadsheet_id,
        protocol_type=request.protocol_type
    )

    try:
        service = get_certification_service(sheets_service)

        result = await service.generate_protocols(
            spreadsheet_id=request.spreadsheet_id,
            protocol_type=request.protocol_type,
            dry_run=request.dry_run
        )

        return CertificationResponse(
            status=result.get("status", "error"),
            rows_affected=0,
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("certification_protocols_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/certification/ds-layouts", response_model=CertificationResponse)
async def generate_ds_layouts(request: CertificationProtocolsRequest):
    """
    Generate DS (Declaration of Safety) layouts.

    NOTE: This endpoint is a placeholder. Full implementation requires
    Google Drive/Docs integration for layout generation.

    Equivalent to GAS generateDsLayouts_353pp().
    """
    from src.services.certification_service import get_certification_service

    logger.info(
        "certification_ds_layouts_requested",
        spreadsheet_id=request.spreadsheet_id
    )

    try:
        service = get_certification_service(sheets_service)

        result = await service.generate_ds_layouts(
            spreadsheet_id=request.spreadsheet_id,
            dry_run=request.dry_run
        )

        return CertificationResponse(
            status=result.get("status", "error"),
            rows_affected=0,
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("certification_ds_layouts_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== Invoice Processing ==============

class InvoiceFormatRequest(BaseModel):
    """Request for invoice formatting."""
    spreadsheet_id: str
    sheet_name: str = "Ордер"
    dry_run: bool = False


class InvoiceCreateRequest(BaseModel):
    """Request for creating full invoice."""
    spreadsheet_id: str
    order_sheet: str = "Ордер"
    certification_sheet: str = "Сертификация"
    labels_sheet: str = "Этикетки"
    target_sheet: str = "Для инвойса"
    dry_run: bool = False


class InvoiceResponse(BaseModel):
    """Response from invoice operations."""
    status: str
    rows_processed: int = 0
    target_sheet: str = ""
    message: str = ""


@api_router.post("/invoice/format-order", response_model=InvoiceResponse)
async def format_order_sheet(request: InvoiceFormatRequest):
    """
    Format order sheet - normalize numeric columns.

    Converts text numbers to actual numbers:
    - Removes spaces, replaces comma with dot
    - Columns: кол-во, Цена ед., Сумма

    Equivalent to GAS formatOrderSheet().
    """
    from src.services.invoice_service import get_invoice_service

    logger.info(
        "invoice_format_requested",
        spreadsheet_id=request.spreadsheet_id,
        sheet_name=request.sheet_name,
        dry_run=request.dry_run
    )

    try:
        service = get_invoice_service(sheets_service)

        result = await service.format_order_sheet(
            spreadsheet_id=request.spreadsheet_id,
            sheet_name=request.sheet_name,
            dry_run=request.dry_run
        )

        return InvoiceResponse(
            status="success" if result.get("status") == "success" else "error",
            rows_processed=result.get("rows_formatted", 0),
            target_sheet=request.sheet_name,
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("invoice_format_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/invoice/create-full", response_model=InvoiceResponse)
async def create_full_invoice(request: InvoiceCreateRequest):
    """
    Create full invoice sheet by joining data from multiple sheets.

    Joins data from:
    - Ордер: base order data (ID, article, quantity, prices)
    - Сертификация: DS names, declarations, spirit info
    - Этикетки: label descriptions

    Creates "Для инвойса" sheet with combined data.

    Equivalent to GAS createFullInvoice().
    """
    from src.services.invoice_service import get_invoice_service

    logger.info(
        "invoice_create_requested",
        spreadsheet_id=request.spreadsheet_id,
        dry_run=request.dry_run
    )

    try:
        service = get_invoice_service(sheets_service)

        result = await service.create_full_invoice(
            spreadsheet_id=request.spreadsheet_id,
            order_sheet=request.order_sheet,
            certification_sheet=request.certification_sheet,
            labels_sheet=request.labels_sheet,
            target_sheet=request.target_sheet,
            dry_run=request.dry_run
        )

        return InvoiceResponse(
            status="success" if result.get("status") == "success" else "error",
            rows_processed=result.get("rows_written", 0),
            target_sheet=result.get("target_sheet", request.target_sheet),
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("invoice_create_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== Formula Operations ==============

class FormulaPriceDynamicsRequest(BaseModel):
    """Request for price dynamics formula recalculation."""
    spreadsheet_id: str
    sheet_name: str = "Динамика цены"
    dry_run: bool = False


class FormulaPriceCalcRequest(BaseModel):
    """Request for price calculation formula update."""
    spreadsheet_id: str
    price_calc_sheet: str = "Расчет цены"
    price_dynamics_sheet: str = "Динамика цены"
    silent: bool = False
    dry_run: bool = False


class FormulaAddYearRequest(BaseModel):
    """Request for adding new year columns."""
    spreadsheet_id: str
    sheet_name: str = "Динамика цены"
    year: Optional[int] = None
    dry_run: bool = False


class FormulaResponse(BaseModel):
    """Response from formula operations."""
    status: str
    blocks_processed: int = 0
    rows_updated: int = 0
    columns_added: int = 0
    year: Optional[int] = None
    message: str = ""


@api_router.post("/formulas/price-dynamics", response_model=FormulaResponse)
async def recalculate_price_dynamics_formulas(request: FormulaPriceDynamicsRequest):
    """
    Recalculate price dynamics formulas.

    Calculates for all year blocks:
    - EXW ALFASPA = EXW * (1 - discount/100)
    - Purchase price = EXW ALFASPA * currency_rate
    - DDP = Purchase * DDP coefficient
    - Growth EXW = current_year / prev_year - 1
    - Growth DDP = current_year_ddp / prev_year_ddp - 1

    Equivalent to GAS recalculatePriceDynamicsFormulas().
    """
    from src.services.formula_service import get_formula_service

    logger.info(
        "formula_price_dynamics_requested",
        spreadsheet_id=request.spreadsheet_id,
        sheet_name=request.sheet_name,
        dry_run=request.dry_run
    )

    try:
        service = get_formula_service(sheets_service)

        result = await service.recalculate_price_dynamics_formulas(
            spreadsheet_id=request.spreadsheet_id,
            sheet_name=request.sheet_name,
            dry_run=request.dry_run
        )

        return FormulaResponse(
            status=result.get("status", "error"),
            blocks_processed=result.get("blocks_processed", 0),
            rows_updated=result.get("rows_updated", 0),
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("formula_price_dynamics_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/formulas/price-calculation", response_model=FormulaResponse)
async def update_price_calculation_formulas(request: FormulaPriceCalcRequest):
    """
    Update price calculation formulas (INDEX/MATCH lookups).

    Pulls data from Price Dynamics sheet based on ID matching:
    - EXW previous year
    - EXW current year
    - EXW ALFASPA current year
    - Purchase price
    - DDP

    Equivalent to GAS updatePriceCalculationFormulas().
    """
    from src.services.formula_service import get_formula_service

    logger.info(
        "formula_price_calc_requested",
        spreadsheet_id=request.spreadsheet_id,
        price_calc_sheet=request.price_calc_sheet,
        dry_run=request.dry_run
    )

    try:
        service = get_formula_service(sheets_service)

        result = await service.update_price_calculation_formulas(
            spreadsheet_id=request.spreadsheet_id,
            price_calc_sheet=request.price_calc_sheet,
            price_dynamics_sheet=request.price_dynamics_sheet,
            silent=request.silent,
            dry_run=request.dry_run
        )

        return FormulaResponse(
            status=result.get("status", "error"),
            rows_updated=result.get("rows_updated", 0),
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("formula_price_calc_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/formulas/add-year-columns", response_model=FormulaResponse)
async def add_new_year_columns(request: FormulaAddYearRequest):
    """
    Add new year columns to price dynamics sheet.

    Inserts 7 columns after "Комментарий":
    - EXW {year}, €
    - СКИДКА ОТ EXW {year}, %
    - EXW ALFASPA {year}, €
    - Закупочная цена {year}, ₽
    - DDP-МОСКВА {year}, ₽
    - Прирост EXW, %
    - Прирост DDP-МОСКВА, %

    Equivalent to GAS addNewYearColumnsToPriceDynamics().
    """
    from src.services.formula_service import get_formula_service

    logger.info(
        "formula_add_year_requested",
        spreadsheet_id=request.spreadsheet_id,
        sheet_name=request.sheet_name,
        year=request.year,
        dry_run=request.dry_run
    )

    try:
        service = get_formula_service(sheets_service)

        result = await service.add_new_year_columns(
            spreadsheet_id=request.spreadsheet_id,
            sheet_name=request.sheet_name,
            year=request.year,
            dry_run=request.dry_run
        )

        return FormulaResponse(
            status=result.get("status", "error"),
            columns_added=result.get("columns_added", 0),
            year=result.get("year"),
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("formula_add_year_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== Log Archiving ==============

class LogArchiveRequest(BaseModel):
    """Request for log archiving."""
    spreadsheet_id: str
    archive_folder_id: str
    project_code: str = "project"
    dry_run: bool = False


class LogResetRequest(BaseModel):
    """Request for log reset."""
    spreadsheet_id: str
    sheet_name: str = "Логи"
    dry_run: bool = False


class LogRotationRequest(BaseModel):
    """Request for midnight log rotation."""
    spreadsheet_id: str
    archive_folder_id: str
    project_code: str = "project"
    force: bool = False
    dry_run: bool = False


class LogEntryRequest(BaseModel):
    """Request for writing log entry."""
    spreadsheet_id: str
    sheet_name: str = "Логи"
    category: str = "SYSTEM"
    action: str
    details: str = ""
    level: str = "INFO"


class LogArchiveResponse(BaseModel):
    """Response from log archiving operations."""
    status: str
    total_rows: int = 0
    sheets_archived: Optional[Dict[str, int]] = None
    archive_name: Optional[str] = None
    message: str = ""


class LogStatusResponse(BaseModel):
    """Response from log status check."""
    status: str
    last_archive_date: str = "Never"
    current_archive_name: str = ""
    project_code: str = ""
    sheets_to_archive: List[str] = []
    current_row_counts: Optional[Dict[str, int]] = None
    total_pending_rows: int = 0


@api_router.post("/logs/archive", response_model=LogArchiveResponse)
async def archive_logs_daily(request: LogArchiveRequest):
    """
    Archive all log sheets to monthly archive spreadsheet.

    Copies data from log sheets (Логи, Журнал синхро, Журнал логов)
    to a monthly archive spreadsheet in the specified Drive folder.

    Equivalent to GAS archiveLogsDaily().
    """
    from src.services.logging_service import get_logging_service

    logger.info(
        "log_archive_requested",
        spreadsheet_id=request.spreadsheet_id,
        project_code=request.project_code,
        dry_run=request.dry_run
    )

    try:
        service = get_logging_service(sheets_service)

        result = await service.archive_logs_daily(
            spreadsheet_id=request.spreadsheet_id,
            archive_folder_id=request.archive_folder_id,
            project_code=request.project_code,
            dry_run=request.dry_run
        )

        return LogArchiveResponse(
            status=result.get("status", "error"),
            total_rows=result.get("total_rows", 0),
            sheets_archived=result.get("sheets_archived"),
            archive_name=result.get("archive_name"),
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("log_archive_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/logs/reset", response_model=LogArchiveResponse)
async def reset_log_sheet(request: LogResetRequest):
    """
    Reset (clear) a log sheet after archiving.

    Clears all data rows, preserving headers.

    Equivalent to GAS resetDailyLogSheet().
    """
    from src.services.logging_service import get_logging_service

    logger.info(
        "log_reset_requested",
        spreadsheet_id=request.spreadsheet_id,
        sheet_name=request.sheet_name,
        dry_run=request.dry_run
    )

    try:
        service = get_logging_service(sheets_service)

        result = await service.reset_daily_log_sheet(
            spreadsheet_id=request.spreadsheet_id,
            sheet_name=request.sheet_name,
            dry_run=request.dry_run
        )

        return LogArchiveResponse(
            status=result.get("status", "error"),
            total_rows=result.get("rows_cleared", 0),
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("log_reset_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/logs/rotation", response_model=LogArchiveResponse)
async def midnight_log_rotation(request: LogRotationRequest):
    """
    Perform complete midnight log rotation.

    1. Archives logs to monthly spreadsheet
    2. Resets all log sheets

    Equivalent to GAS midnightLogRotation().
    """
    from src.services.logging_service import get_logging_service

    logger.info(
        "log_rotation_requested",
        spreadsheet_id=request.spreadsheet_id,
        project_code=request.project_code,
        force=request.force,
        dry_run=request.dry_run
    )

    try:
        service = get_logging_service(sheets_service)

        result = await service.midnight_log_rotation(
            spreadsheet_id=request.spreadsheet_id,
            archive_folder_id=request.archive_folder_id,
            project_code=request.project_code,
            force=request.force,
            dry_run=request.dry_run
        )

        archive_result = result.get("archive_result", {})
        return LogArchiveResponse(
            status=result.get("status", "error"),
            total_rows=archive_result.get("total_rows", 0),
            sheets_archived=archive_result.get("sheets_archived"),
            archive_name=archive_result.get("archive_name"),
            message=result.get("message", "")
        )

    except Exception as e:
        logger.error("log_rotation_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.get("/logs/status", response_model=LogStatusResponse)
async def get_log_status(spreadsheet_id: str, project_code: str = "project"):
    """
    Get current log archive status.

    Returns last archive date, pending row counts, and configuration.

    Equivalent to GAS showArchiveStatus().
    """
    from src.services.logging_service import get_logging_service

    logger.info(
        "log_status_requested",
        spreadsheet_id=spreadsheet_id,
        project_code=project_code
    )

    try:
        service = get_logging_service(sheets_service)

        result = await service.get_archive_status(
            spreadsheet_id=spreadsheet_id,
            project_code=project_code
        )

        return LogStatusResponse(
            status=result.get("status", "error"),
            last_archive_date=result.get("last_archive_date", "Never"),
            current_archive_name=result.get("current_archive_name", ""),
            project_code=result.get("project_code", project_code),
            sheets_to_archive=result.get("sheets_to_archive", []),
            current_row_counts=result.get("current_row_counts"),
            total_pending_rows=result.get("total_pending_rows", 0)
        )

    except Exception as e:
        logger.error("log_status_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/logs/write")
async def write_log_entry(request: LogEntryRequest):
    """
    Write a log entry to the log sheet.

    Equivalent to GAS _logToSheet_().
    """
    from src.services.logging_service import get_logging_service

    try:
        service = get_logging_service(sheets_service)

        result = await service.write_log_entry(
            spreadsheet_id=request.spreadsheet_id,
            sheet_name=request.sheet_name,
            category=request.category,
            action=request.action,
            details=request.details,
            level=request.level
        )

        return result

    except Exception as e:
        logger.error("log_write_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== AI Analysis ==============

@api_router.get("/ai/status")
async def check_ai_service():
    """
    Check Gemini AI service status.
    Returns connection status, current model, and available models.
    """
    logger.info("ai_status_check_requested")

    try:
        result = ai_service.check_service()
        return {
            "status": result.get("status"),
            "model": result.get("model"),
            "response": result.get("response"),
            "available_models": result.get("available_models", []),
            "error": result.get("error")
        }
    except Exception as e:
        logger.error("ai_status_check_failed", error=str(e))
        return {
            "status": "error",
            "error": str(e)
        }


@api_router.get("/ai/settings")
async def get_ai_settings():
    """
    Get current Gemini AI settings.
    Returns model, API key status (masked), and config.
    """
    logger.info("ai_settings_requested")

    try:
        settings = ai_service.get_gemini_settings()
        return settings
    except Exception as e:
        logger.error("ai_settings_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


class AIConfigureRequest(BaseModel):
    """Request to configure AI settings."""
    api_key: Optional[str] = None
    model: Optional[str] = None


@api_router.post("/ai/configure")
async def configure_ai(request: AIConfigureRequest):
    """
    Configure Gemini AI settings at runtime.
    Allows setting API key and model without server restart.

    Note: Settings are stored in memory and will be lost on server restart.
    For permanent settings, update .env file.
    """
    from src.services.gemini_client import set_runtime_config, get_gemini_client

    logger.info("ai_configure_requested",
               has_key=bool(request.api_key),
               model=request.model)

    try:
        # Update runtime config
        set_runtime_config(
            api_key=request.api_key,
            model=request.model
        )

        # Test the new configuration
        client = get_gemini_client()
        test_result = client.test_connection()

        if test_result.get("status") == "online":
            return {
                "status": "success",
                "message": "Configuration updated successfully",
                "model": client.model,
                "test_response": test_result.get("response"),
                "available_models": test_result.get("available_models", [])[:5]
            }
        else:
            return {
                "status": "error",
                "message": "Configuration updated but connection test failed",
                "error": test_result.get("error"),
                "model": client.model
            }

    except Exception as e:
        logger.error("ai_configure_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.get("/ai/categories")
async def get_ai_categories():
    """
    Get available ТН ВЭД categories for classification.
    """
    return {
        "categories": ai_service.get_available_categories()
    }


@api_router.post("/ai/analyze/pdf")
async def analyze_pdf(request: AIPdfAnalyzeRequest):
    """
    Analyze a PDF document directly without spreadsheet.
    Extracts INCI composition and classifies product.
    """
    logger.info("pdf_analysis_requested", url=request.pdf_url[:50] if request.pdf_url else None)

    try:
        result = ai_service.analyze_pdf(
            pdf_url=request.pdf_url,
            purpose=request.purpose,
            application=request.application
        )
        return {
            "status": "success",
            "data": result
        }
    except Exception as e:
        logger.error("pdf_analysis_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/ai/analyze/simple")
async def analyze_simple(request: AISimpleAnalyzeRequest):
    """
    Simple AI analysis without PDF - just classify product by name/purpose.
    Use this to quickly test if Gemini API is working.
    Returns product_type and product_reason.
    """
    import json
    from src.services.gemini_client import get_gemini_client

    logger.info("simple_analysis_requested", 
                product=request.product_name[:50] if request.product_name else None)

    try:
        # Build prompt for simple classification
        inci_part = f"\nИНСИ состав: {request.inci_text}" if request.inci_text else ""
        
        prompt = f"""Ты эксперт по косметической продукции.

Продукт: {request.product_name}
Назначение: {request.purpose or 'не указано'}
Применение: {request.application or 'не указано'}{inci_part}

Определи:
1. Вид продукции (например: крем для лица, сыворотка, шампунь, маска и т.д.)
2. Обоснование почему это именно такой вид продукции

Ответ ТОЛЬКО в формате JSON:
{{"product_type": "Вид продукции", "product_reason": "Обоснование"}}"""

        client = get_gemini_client()
        response = client.call_api(prompt, json_response=True, temperature=0.1)
        
        # Parse JSON response
        cleaned = response.strip()
        if cleaned.startswith("```json"):
            cleaned = cleaned[7:]
        if cleaned.startswith("```"):
            cleaned = cleaned[3:]
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3]
        
        result = json.loads(cleaned.strip())
        
        # Handle case when Gemini returns a list instead of dict
        if isinstance(result, list):
            if len(result) > 0 and isinstance(result[0], dict):
                result = result[0]
            else:
                raise ValueError(f"Unexpected list response from Gemini: {result}")
        
        logger.info("simple_analysis_success",
                   product_type=result.get("product_type"))

        return {
            "status": "success",
            "data": result
        }
    except Exception as e:
        logger.error("simple_analysis_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/ai/analyze/row")
async def analyze_row(request: AIAnalyzeRequest):
    """
    Analyze a single row in the spreadsheet.
    Reads INCI link from column H, analyzes with Gemini,
    and writes results to columns L-Y.
    """
    if not request.row_number:
        raise HTTPException(status_code=400, detail="row_number is required")

    logger.info(
        "row_analysis_requested",
        spreadsheet_id=request.spreadsheet_id[:20],
        sheet=request.sheet_name,
        row=request.row_number
    )

    try:
        result = ai_service.analyze_row(
            spreadsheet_id=request.spreadsheet_id,
            sheet_name=request.sheet_name,
            row_number=request.row_number
        )

        if result.success:
            logging_service.add_log(
                request.spreadsheet_id,
                "AI АНАЛИЗ",
                f"Анализ строки {request.row_number}",
                f"Категория: {result.data.get('category_code', 'N/A')}, "
                f"Вид: {result.data.get('product_type', 'N/A')}",
                "✅ УСПЕХ"
            )
            return {
                "status": "success",
                "row_number": result.row_number,
                "duration": f"{result.duration:.2f}s",
                "data": result.data
            }
        else:
            logging_service.add_log(
                request.spreadsheet_id,
                "AI АНАЛИЗ",
                f"Ошибка анализа строки {request.row_number}",
                result.error or "Unknown error",
                "❌ ОШИБКА"
            )
            raise HTTPException(status_code=500, detail=result.error)

    except HTTPException:
        raise
    except Exception as e:
        logger.error("row_analysis_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.post("/ai/analyze/batch")
async def analyze_batch(request: AIAnalyzeBatchRequest, background_tasks: BackgroundTasks):
    """
    Analyze all empty rows in the spreadsheet (batch mode).
    Finds rows with INCI link but no analysis results and processes them.
    Runs in background.
    """
    logger.info(
        "batch_analysis_requested",
        spreadsheet_id=request.spreadsheet_id[:20],
        sheet=request.sheet_name
    )

    try:
        # Run analysis in background
        background_tasks.add_task(
            _run_batch_analysis,
            request.spreadsheet_id,
            request.sheet_name,
            request.delay_between
        )

        logging_service.add_log(
            request.spreadsheet_id,
            "AI АНАЛИЗ",
            "Запущен пакетный анализ",
            f"Лист: {request.sheet_name}",
            "📥 В ОЧЕРЕДИ"
        )

        return {
            "status": "queued",
            "message": "Batch analysis started in background",
            "sheet_name": request.sheet_name
        }

    except Exception as e:
        logger.error("batch_analysis_start_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


async def _run_batch_analysis(spreadsheet_id: str, sheet_name: str, delay: float):
    """Background task for batch analysis."""
    try:
        result = ai_service.analyze_empty_rows(
            spreadsheet_id=spreadsheet_id,
            sheet_name=sheet_name,
            delay_between=delay
        )

        logging_service.add_log(
            spreadsheet_id,
            "AI АНАЛИЗ",
            "Пакетный анализ завершен",
            f"Успешно: {result['success']}/{result['total']}, Ошибок: {result['failed']}",
            "✅ ЗАВЕРШЕН" if result['failed'] == 0 else "⚠️ С ОШИБКАМИ"
        )

        logger.info("batch_analysis_completed",
                   total=result['total'],
                   success=result['success'],
                   failed=result['failed'])

    except Exception as e:
        logger.error("batch_analysis_background_failed", error=str(e))
        logging_service.add_log(
            spreadsheet_id,
            "AI АНАЛИЗ",
            "Ошибка пакетного анализа",
            str(e),
            "❌ ОШИБКА"
        )


# ============== Cache Management ==============

@api_router.post("/cache/clear")
async def clear_cache(project: Optional[str] = None, pattern: Optional[str] = None):
    """Clear cache."""
    logger.info("cache_clear_requested", project=project, pattern=pattern)

    # TODO: Clear cache
    return {"cleared_keys": 0, "message": "Cache cleared"}


# ============== Task Status ==============

@api_router.get("/status/{task_id}", response_model=TaskStatusResponse)
async def get_task_status(task_id: str):
    """Get status of any task."""
    # TODO: Get from task queue
    return TaskStatusResponse(
        task_id=task_id,
        status="pending",
    )


# ============== Menu Configuration ==============

class MenuItemModel(BaseModel):
    """Single menu item."""
    label: Optional[str] = None
    function_name: Optional[str] = None
    separator: bool = False
    separator_after: bool = False
    submenu: Optional[str] = None
    items: Optional[List[dict]] = None


class MenuGroupModel(BaseModel):
    """Menu group with items."""
    title: str
    items: List[MenuItemModel]

class MenuConfigResponse(BaseModel):
    """Menu configuration for a project."""
    project: str
    project_name: str
    menu_title: str
    items: List[MenuItemModel] = []
    menus: List[MenuGroupModel] = []

# Project spreadsheet IDs mapping (matches GAS 01Config.js DOC_TO_PROJECT)
PROJECT_IDS = {
    # MT (Montibello)
    "13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ": "MT", # Main
    "1fMOjUE7oZV96fCY5j5rPxnhWGJkDqg-GfwPZ8jUVgPw": "MT", # Alt Main
    "1BW8Gk5_X2EZVjbnaa2yDm-bPzzlggwQrHepeNCcPCc0": "MT", # Source
    "199Np7xsBiBRQih5_tlUdpt6EmkfRGjZAhTvKm4Ua0Q6XEaMtvAmQUn0g": "MT", # Legacy?
    
    # SS (San)
    "12yIL1CuESZxeUUd-oKK2brtN1FnXE9q95N7SqzNc7vk": "SS", # Main
    "1sTgZa-n1aP7oIhyQfPeN8QDgDNnCubqMWAd-TKjKpJXWsQm_ZhXnojPD": "SS", # Source
    "1J8Yzfz9621gqJkPh5ZKBa0v34nv3v9_7OL4JIROlHj0": "SS", # User Source

    # SK (Carmado)
    "1CpYYLvRYslsyCkuLzL9EbbjsvbNpWCEZcmhKqMoX5zw": "SK", # Main
    "1zSu0PzKKa5wvwMZCicwLN8N7Rwhs8XlJVrTrt2LMzQs": "SK", # Source
    "1DJvK1vUT2OTubN0TLdZvsgYMSYByLHl8xTsus3K-KJ-VtJxgGnSw5Ih8": "SK", # Legacy?
}

PROJECT_NAMES = {
    "MT": "CosmeticaBar (MT)",
    "SK": "Carmado (SK)",
    "SS": "San (SS)",
}

# Menu configurations per project
MENU_CONFIGS = {
    "MT": {
        "menu_title": "MT CosmeticaBar",
        "order_sheet": "Заказ",
        "sort_columns": {
            "manufacturer": "Производитель",
            "price": "EXW ALFASPA текущая, €",
        },
        "primary_menu": {
            "title": "🧾 Заказ",
            "items": {
                "MAIN": "Обработка Б/З поставщик",
                "TESTER": "Обработка Тестер",
                "SAMPLES": "Обработка Пробники",
                "STOCKS": "Загрузить остатки",
                "NEW_PRICE_YEAR": "New год для динамика",
            },
        },
        "order_stages_menu": {
            "title": "📊 Стадии по заказ",
            "items": {
                "SORT_MANUFACTURER": "Сортировать по производителю",
                "SORT_PRICE": "Сортировать по прайсу",
                "STAGE_ALL": "1. Все данные",
                "STAGE_ORDER": "2. Заказ",
                "STAGE_PROMOTIONS": "3. Акции",
                "STAGE_SET": "4. Набор",
                "STAGE_PRICE": "5. Прайс",
            },
        },
    },
    "SK": {
        "menu_title": "SK Carmado",
        "order_sheet": "Заказ",
        "sort_columns": {
            "manufacturer": "Производитель",
            "price": "Цена",
        },
        "primary_menu": {
            "title": "🧾 Заказ",
            "items": {
                "MAIN": "Обработка Б/З поставщик",
                "PROBES": "Обработка пробники",
                "STOCKS": "Загрузить остатки",
                "NEW_PRICE_YEAR": "New год для динамика",
            },
        },
        "order_stages_menu": {
            "title": "📊 Стадии по заказ",
            "items": {
                "SORT_MANUFACTURER": "Сортировать по производителю",
                "SORT_PRICE": "Сортировать по прайсу",
                "STAGE_ALL": "1. Все данные",
                "STAGE_ORDER": "2. Заказ",
                "STAGE_PROMOTIONS": "3. Акции",
                "STAGE_SET": "4. Набор",
                "STAGE_PRICE": "5. Прайс",
            },
        },
    },
    "SS": {
        "menu_title": "SS San",
        "order_sheet": "Заказ",
        "sort_columns": {
            "manufacturer": "Производитель",
            "price": "Цена",
        },
        "primary_menu": {
            "title": "🧾 Заказ",
            "items": {
                "MAIN": "Обработка Б/З поставщик",
                "STOCKS": "Загрузить остатки",
                "NEW_PRICE_YEAR": "New год для динамика",
            },
        },
        "order_stages_menu": {
            "title": "📊 Стадии по заказ",
            "items": {
                "SORT_MANUFACTURER": "Сортировать по производителю",
                "SORT_PRICE": "Сортировать по прайсу",
                "STAGE_ALL": "1. Все данные",
                "STAGE_ORDER": "2. Заказ",
                "STAGE_PROMOTIONS": "3. Акции",
                "STAGE_SET": "4. Набор",
                "STAGE_PRICE": "5. Прайс",
            },
        },
    },
}

PRIMARY_DATA_MENU_ORDER = ["MAIN", "TESTER", "SAMPLES", "PROBES", "STOCKS", "NEW_PRICE_YEAR"]
ORDER_STAGES_MENU_ORDER = [
    "SORT_MANUFACTURER",
    "SORT_PRICE",
    "STAGE_ALL",
    "STAGE_ORDER",
    "STAGE_PROMOTIONS",
    "STAGE_SET",
    "STAGE_PRICE",
]

# Action mapping for menu buttons (mirrors legacy GAS config)
PRIMARY_DATA_MENU_ACTIONS = {
    "MAIN": {"fn_by_project": {"SK": "serverProcessSkMain", "SS": "serverProcessSsMain", "MT": "serverProcessMtMain"}},
    "TESTER": {"fn_by_project": {"MT": "serverProcessMtTester"}},
    "SAMPLES": {"fn_by_project": {"MT": "serverProcessMtSamples"}},
    "PROBES": {"fn_by_project": {"SK": "serverProcessSkProbes"}},
    "STOCKS": {"fn_by_project": {"SK": "loadSkStockData", "SS": "loadSsStockData", "MT": "loadMtStockData"}},
    "NEW_PRICE_YEAR": {"fn": "serverAddNewYearColumns"},
    "SORT_MANUFACTURER": {"fn": "sortByManufacturer"},
    "SORT_PRICE": {"fn": "sortByPrice"},
    "STAGE_ALL": {"fn": "serverShowAllOrderData"},
    "STAGE_ORDER": {"fn": "serverShowOrderStage"},
    "STAGE_PROMOTIONS": {"fn": "serverShowPromotionsStage"},
    "STAGE_SET": {"fn": "serverShowSetStage"},
    "STAGE_PRICE": {"fn": "serverShowPriceStage"},
}

# Static menu groups from legacy GAS config (without project-specific items)
BASE_MENU_GROUPS: List[dict] = [
    {
        "title": "⚙️ ЭКОСИСТЕМА",
        "items": [
            {
                "submenu": "🔍 Проверки систем",
                "items": [
                    {"label": "📊 Показать статус всех сервисов", "function_name": "showAllServicesStatus_proxy"},
                ],
            },
            {
                "submenu": "🤖 Агент",
                "items": [
                    {"label": "Проверить сервис агента", "function_name": "menuCheckService"},
                    {"label": "⚙️ Настройки Gemini", "function_name": "setupGeminiComplete"},
                    {"label": "📋 Показать текущие настройки", "function_name": "showGeminiSettings"},
                ],
            },
            {
                "submenu": "📊 Логи",
                "items": [
                    {"label": "📈 Дашборд логов", "function_name": "openLogDashboard_proxy"},
                    {"separator": True},
                    {"label": "📑 Показать журнал синхро", "function_name": "showSyncJournal_proxy"},
                    {"label": "📑 Показать журнал логов", "function_name": "showLogJournal_proxy"},
                    {"separator": True},
                    {"label": "📦 Архивировать логи", "function_name": "manualArchiveLogs_proxy"},
                    {"label": "📊 Статус архива", "function_name": "showArchiveStatus_proxy"},
                    {"separator": True},
                    {"label": "🔄 Пересоздать журнал синхро", "function_name": "recreateLogSheet"},
                    {"label": "🔄 Пересоздать журнал логов", "function_name": "recreateDebugLogSheet"},
                    {"label": "🧹 Очистить журнал (быстро)", "function_name": "quickCleanLogSheet"},
                ],
            },
            {"separator": True},
            {"label": "➕ Добавить артикул", "function_name": "addArticleManually"},
            {"label": "❌ Удалить артикул", "function_name": "deleteSelectedRowsWithSync"},
            {"label": "🔄 Синхронизировать строку", "function_name": "syncSelectedRow"},
            {"label": "🔄 Синхронизировать всю таблицу", "function_name": "runFullSync"},
        ],
    },
    {
        "title": "📦 Выгрузка",
        "items": [
            {"label": "Выгрузить Акции", "function_name": "serverExportPromotions"},
            {"label": "Выгрузить Наборы", "function_name": "serverExportSets"},
        ],
    },
    {
        "title": "🚚 Поставка",
        "items": [
            {"label": "Форматировать лист 'Ордер'", "function_name": "serverFormatOrderSheet"},
            {"separator": True},
            {"label": "1. Создать лист 'Для инвойса'", "function_name": "serverCreateFullInvoice"},
            {"label": "2. Собрать документы", "function_name": "collectAndCopyDocuments_proxy"},
        ],
    },
    {
        "title": "🔬 Сертификация",
        "items": [
            {"label": "Лист новинки", "function_name": "serverCreateNewsSheet"},
            {"separator": True},
            {"label": "Создать заявку протоколы (353пп)", "function_name": "serverGenerateProtocols353pp"},
            {"label": "Создать заявку ДС (353пп)", "function_name": "serverGenerateDsLayouts"},
            {"label": "Собрать документы для заявки (353пп)", "function_name": "structureDocuments_353pp"},
            {"separator": True},
            {"label": "Посчитать спирты", "function_name": "serverCalculateSpiritNumbers"},
            {"label": "Создать Макеты спирты", "function_name": "generateSpiritProtocols"},
            {"separator": True},
            {"label": "Пересчитать каскады (Сертификация)", "function_name": "serverRecalculateCascades"},
        ],
    },
    {
        "title": "⚙️ СИНХРОНИЗАЦИЯ",
        "items": [
            {"label": "📝 Настроить правила", "function_name": "showSyncRulesManagerDialog"},
            {"label": "🧹 Очистить уведомления", "function_name": "clearAllToasts"},
            {"label": "🚀 Миграция правил", "function_name": "migrateLegacyRules"},
            {"label": "🗑️ Удалить лист 'Правила синхро'", "function_name": "deleteLegacyRulesSheet"},
            {"separator": True},
            {
                "submenu": "Операции с артикулами",
                "items": [
                    {"label": "Добавить артикул", "function_name": "addArticleManually"},
                    {"label": "Удалить артикул", "function_name": "deleteSelectedRowsWithSync"},
                    {"label": "Синхронизировать строку", "function_name": "syncSelectedRow"},
                    {"label": "Синхронизировать ВСЮ таблицу", "function_name": "runFullSync"},
                ],
            },
            {"separator": True},
            {"label": "🔄 Обновить триггеры", "function_name": "setupTriggers"},
        ],
    },
    {
        "title": "🤖 АГЕНТ",
        "items": [
            {"label": "🎯 Smart Match для строки", "function_name": "menuSmartMatch"},
            {"label": "🤖 Анализировать выбранную строку", "function_name": "menuAnalyzeSelected"},
            {"label": "📊 Анализировать пустые строки", "function_name": "menuAnalyzeEmpty"},
            {"separator": True},
            {"label": "📦 Показать категории ТН ВЭД", "function_name": "menuShowCategories"},
        ],
    },
]


def _resolve_action_fn(action_def: dict, project: str) -> Optional[str]:
    """Pick correct function name for menu action."""
    if not action_def:
        return None
    if "fn_by_project" in action_def:
        return action_def["fn_by_project"].get(project)
    return action_def.get("fn")


def _build_primary_menu(project: str) -> Optional[MenuGroupModel]:
    config = MENU_CONFIGS.get(project, MENU_CONFIGS["MT"])
    menu_cfg = config.get("primary_menu")
    if not menu_cfg:
        return None

    items_cfg = menu_cfg.get("items", {})
    items: List[MenuItemModel] = []

    for key in PRIMARY_DATA_MENU_ORDER:
        label = items_cfg.get(key)
        if not label:
            continue

        action_def = PRIMARY_DATA_MENU_ACTIONS.get(key)
        fn = _resolve_action_fn(action_def, project)
        if not fn:
            continue

        items.append(MenuItemModel(label=label, function_name=fn))

    if not items:
        return None

    return MenuGroupModel(title=menu_cfg.get("title", "🧾 Заказ"), items=items)


def _build_order_stages_menu(project: str) -> Optional[MenuGroupModel]:
    config = MENU_CONFIGS.get(project, MENU_CONFIGS["MT"])
    stages_cfg = config.get("order_stages_menu")
    if not stages_cfg:
        return None

    items_cfg = stages_cfg.get("items", {})
    items: List[MenuItemModel] = []

    for key in ORDER_STAGES_MENU_ORDER:
        label = items_cfg.get(key)
        if not label:
            continue

        action_def = PRIMARY_DATA_MENU_ACTIONS.get(key)
        fn = _resolve_action_fn(action_def, project)
        if not fn:
            continue

        items.append(MenuItemModel(label=label, function_name=fn))

    if not items:
        return None

    return MenuGroupModel(title=stages_cfg.get("title", "📊 Стадии по заказ"), items=items)


def _clone_base_group(group: dict) -> MenuGroupModel:
    cloned_items = []
    for item in group.get("items", []):
        cloned_items.append(MenuItemModel(**item))
    return MenuGroupModel(title=group.get("title", "Меню"), items=cloned_items)


def _server_tools_menu() -> MenuGroupModel:
    """Utility menu for server-driven actions and diagnostics."""
    items = [
        MenuItemModel(label="🔄 Обновить меню", function_name="refreshMenu"),
        MenuItemModel(label="📑 Упорядочить листы", function_name="reorderSheets"),
        MenuItemModel(label="🔄 Обновить данные", function_name="callServerLoadFunctions"),
        MenuItemModel(label="🟢 Статус сервера", function_name="checkServerStatus"),
        MenuItemModel(label="🐛 Debug: Spreadsheet ID", function_name="debugShowSpreadsheetId"),
    ]
    return MenuGroupModel(title="🟢 Ecosystem", items=items)


def _build_menu_registry(project: str) -> List[MenuGroupModel]:
    """Assemble full menu registry similar to legacy GAS config."""
    registry: List[MenuGroupModel] = []
    
    # 1. Извлекаем ЭКОСИСТЕМУ из базовых групп и ставим первой
    base_groups = [_clone_base_group(g) for g in BASE_MENU_GROUPS]
    ecosystem_group = next((g for g in base_groups if "ЭКОСИСТЕМА" in g.title), None)
    if ecosystem_group:
        registry.append(ecosystem_group)
        base_groups = [g for g in base_groups if g != ecosystem_group]

    # 2. Добавляем динамические группы (Заказ, Стадии)
    primary_group = _build_primary_menu(project)
    if primary_group:
        registry.append(primary_group)

    stages_group = _build_order_stages_menu(project)
    if stages_group:
        registry.append(stages_group)

    # 3. Добавляем остальные базовые группы
    registry.extend(base_groups)

    # 4. Диагностика сервера
    registry.append(_server_tools_menu())

    return registry

@api_router.post("/logs/init")
async def init_logs(request: LogInitRequest):
    """Initialize the 'Логи' sheet."""
    logger.info("logs_init_requested", spreadsheet_id=request.spreadsheet_id)
    success = logging_service.init_session_log(request.spreadsheet_id)
    if not success:
        raise HTTPException(status_code=500, detail="Failed to initialize logs sheet")
    return {"status": "success", "message": "Log sheet initialized"}


@api_router.get("/menu/config")
async def get_menu_config(spreadsheet_id: str) -> MenuConfigResponse:
    """Get menu configuration for a spreadsheet."""
    project = PROJECT_IDS.get(spreadsheet_id, "MT")
    config = MENU_CONFIGS.get(project, MENU_CONFIGS["MT"])

    logger.info("menu_config_requested", spreadsheet_id=spreadsheet_id, project=project)

    registry = _build_menu_registry(project)
    first_group_items = registry[0].items if registry else []

    return MenuConfigResponse(
        project=project,
        project_name=PROJECT_NAMES.get(project, project),
        menu_title=config["menu_title"],
        items=first_group_items,
        menus=registry,
    )

@api_router.get("/menu/sort-config")
async def get_sort_config(spreadsheet_id: str):
    """Get sort configuration for a spreadsheet (columns, sheet names)."""
    project = PROJECT_IDS.get(spreadsheet_id, "MT")
    config = MENU_CONFIGS.get(project, MENU_CONFIGS["MT"])

    return {
        "project": project,
        "order_sheet": config["order_sheet"],
        "sort_columns": config["sort_columns"]
    }


# ============== Sheet Ordering ==============

# Standard sheet order for all projects (matches GAS Lib.reorderSheets)
SHEET_ORDER = [
    "-Б/З поставщик",
    "-Тестер",
    "-Пробники",
    "-пробники",
    "-остатки",
    "Главная",
    "Заказ",
    "Динамика цены",
    "Расчет цены",
    "Прайс",
    "Этикетки",
    "Сертификация",
    "ABC-Анализ",
    "ТЗ по статусам",
    "Сверка заказа",
    "Для таможни",
    "New sert",
    "Для базы",
    # Auxiliary sheets at the end
    "Вид и код",
    "Справочник",
    "Журнал синхро",
    "Правила синхро",
    "Информация",
]

class ReorderSheetsRequest(BaseModel):
    """Request to reorder sheets."""
    spreadsheet_id: str

@api_router.post("/sheets/reorder")
async def reorder_sheets(request: ReorderSheetsRequest):
    """
    Reorder sheets according to standard order.
    Unknown sheets are moved to the end.
    """
    logger.info("reorder_sheets_requested", spreadsheet_id=request.spreadsheet_id)

    try:
        result = sheets_service.reorder_sheets(request.spreadsheet_id, SHEET_ORDER)
        return {
            "status": "success",
            "message": f"Reordered {result['sheets_moved']} sheets",
            "details": result
        }
    except Exception as e:
        logger.error("reorder_sheets_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== SYNC LOG ENDPOINTS ==============


class SyncLogQueryParams(BaseModel):
    spreadsheet_id: str
    limit: int = 200
    offset: int = 0
    order: str = "desc"
    project: Optional[str] = None
    category: Optional[str] = None
    status: Optional[str] = None
    row_key: Optional[str] = None
    source: Optional[str] = None
    target: Optional[str] = None
    rule_id: Optional[str] = None


@api_router.get("/sync-logs")
def get_sync_logs(
    spreadsheet_id: str,
    limit: int = 200,
    offset: int = 0,
    order: str = "desc",
    project: Optional[str] = None,
    category: Optional[str] = None,
    status: Optional[str] = None,
    row_key: Optional[str] = None,
    source: Optional[str] = None,
    target: Optional[str] = None,
    rule_id: Optional[str] = None,
):
    """
    Get sync journal entries with filtering and pagination.
    """
    try:
        entries, total = sync_log_service.list_entries(
            spreadsheet_id=spreadsheet_id,
            limit=limit,
            offset=offset,
            order=order,
            project=project,
            category=category,
            status=status,
            row_key=row_key,
            source=source,
            target=target,
            rule_id=rule_id,
        )
        return {
            "status": "success",
            "total": total,
            "limit": limit,
            "offset": offset,
            "entries": entries,
        }
    except Exception as e:
        logger.error("get_sync_logs_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.get("/sync-logs/stats")
def get_sync_logs_stats(spreadsheet_id: str):
    """
    Get summary statistics for sync logs.
    """
    try:
        entries, total = sync_log_service.list_entries(
            spreadsheet_id=spreadsheet_id,
            limit=10000,
            offset=0,
        )
        
        # Calculate stats
        categories = {}
        statuses = {}
        for entry in entries:
            cat = entry.get("category", "unknown") or "unknown"
            categories[cat] = categories.get(cat, 0) + 1
            
            st = entry.get("status", "unknown") or "unknown"
            statuses[st] = statuses.get(st, 0) + 1
        
        return {
            "status": "success",
            "total_entries": total,
            "categories": categories,
            "statuses": statuses,
        }
    except Exception as e:
        logger.error("get_sync_logs_stats_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# Path-based versions for UI compatibility (logs_manager.html)

@api_router.get("/sync-logs/{spreadsheet_id}")
def get_sync_logs_by_path(
    spreadsheet_id: str,
    limit: int = 200,
    offset: int = 0,
    order: str = "desc",
    project: Optional[str] = None,
    category: Optional[str] = None,
    status: Optional[str] = None,
):
    """
    Get sync journal entries (path-based for UI compatibility).
    """
    try:
        entries, total = sync_log_service.list_entries(
            spreadsheet_id=spreadsheet_id,
            limit=limit,
            offset=offset,
            order=order,
            project=project,
            category=category,
            status=status,
        )
        return {
            "items": entries,
            "total": total,
            "limit": limit,
            "offset": offset,
        }
    except Exception as e:
        logger.error("get_sync_logs_by_path_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.get("/sync-logs/{spreadsheet_id}/stats")
def get_sync_logs_stats_by_path(spreadsheet_id: str):
    """
    Get sync log statistics (path-based for UI compatibility).
    """
    try:
        entries, total = sync_log_service.list_entries(
            spreadsheet_id=spreadsheet_id,
            limit=10000,
            offset=0,
        )
        
        by_category = {}
        by_status = {}
        for entry in entries:
            cat = entry.get("category", "unknown") or "unknown"
            by_category[cat] = by_category.get(cat, 0) + 1
            
            st = entry.get("status", "unknown") or "unknown"
            by_status[st] = by_status.get(st, 0) + 1
        
        return {
            "total": total,
            "by_category": by_category,
            "by_status": by_status,
        }
    except Exception as e:
        logger.error("get_sync_logs_stats_by_path_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== FUNCTION LOG ENDPOINTS ==============


@api_router.get("/function-logs")
def get_function_logs(
    module: Optional[str] = None,
    function_name: Optional[str] = None,
    limit: int = 100,
    offset: int = 0,
):
    """
    Get function execution logs with filtering.
    """
    try:
        executions, total = function_log_service.list_executions(
            module=module,
            function_name=function_name,
            limit=limit,
            offset=offset,
        )
        return {
            "status": "success",
            "total": total,
            "limit": limit,
            "offset": offset,
            "executions": executions,
        }
    except Exception as e:
        logger.error("get_function_logs_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


@api_router.get("/function-logs/{execution_id}")
def get_function_execution(execution_id: str):
    """
    Get a specific function execution by ID.
    """
    try:
        execution = function_log_service.get_execution(execution_id)
        if execution:
            return {"status": "success", "execution": execution}
        
        # Try to find in stored executions
        executions, _ = function_log_service.list_executions(limit=1000)
        for exec_data in executions:
            if exec_data.get("execution_id") == execution_id:
                return {"status": "success", "execution": exec_data}
        
        raise HTTPException(status_code=404, detail="Execution not found")
    except HTTPException:
        raise
    except Exception as e:
        logger.error("get_function_execution_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))
