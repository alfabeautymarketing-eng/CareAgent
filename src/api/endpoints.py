"""
API endpoints for sync, price processing, AI analysis.
"""

from datetime import datetime
from typing import Optional, List

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from src.utils.logger import logger

api_router = APIRouter()

from src.services.sheets import SheetsService

sheets_service = SheetsService()


# ============== Request/Response Models ==============

class SyncRowRequest(BaseModel):
    """Request to sync a single row."""

    project: str
    article: str
    source_sheet: str
    target_sheets: Optional[List[str]] = None


class SyncRangeRequest(BaseModel):
    """Request to sync a range."""

    project: str
    source_sheet: str
    range: str


class PriceProcessRequest(BaseModel):
    """Request to process price file."""

    project: str
    file_url: Optional[str] = None
    file_id: Optional[str] = None
    start_phase: int = 1
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

    project: str
    row_number: Optional[int] = None
    article: Optional[str] = None
    pdf_url: Optional[str] = None
    product_name: Optional[str] = None


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

    # TODO: Implement sync logic
    # result = await sync_service.sync_row(request)

    return {
        "task_id": f"row_{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "status": "accepted",
        "message": f"Row sync queued for {request.article}",
    }


@api_router.post("/sync/range")
async def sync_range(request: SyncRangeRequest):
    """Sync a range of cells."""
    logger.info(
        "sync_range_requested",
        project=request.project,
        sheet=request.source_sheet,
        range=request.range,
    )

    # TODO: Implement sync logic
    return {
        "task_id": f"range_{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "status": "accepted",
    }


@api_router.post("/sync/full")
async def sync_full(project: str, source_sheet: str):
    """Full sync of a sheet."""
    logger.info("full_sync_requested", project=project, sheet=source_sheet)

    # TODO: Implement full sync
    return {
        "task_id": f"full_{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "status": "accepted",
    }


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
        sheets_service.sort_by_header(
            request.spreadsheet_id, 
            request.sheet_name, 
            request.column_name, 
            request.ascending
        )
        return {"status": "success", "message": f"Sorted by {request.column_name}"}
    except Exception as e:
        logger.error("sort_endpoint_error", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))


# ============== Price Processing ==============

@api_router.post("/price/process")
async def process_price(request: PriceProcessRequest):
    """Process price file (9 phases)."""
    logger.info(
        "price_process_requested",
        project=request.project,
        start_phase=request.start_phase,
    )

    # TODO: Implement price processing
    return {
        "task_id": f"price_{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "status": "accepted",
        "estimated_duration_seconds": 120,
    }


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


# ============== AI Analysis ==============

@api_router.post("/ai/analyze/inci")
async def analyze_inci(request: AIAnalyzeRequest):
    """Analyze INCI composition."""
    logger.info(
        "inci_analysis_requested",
        project=request.project,
        row=request.row_number,
        article=request.article,
    )

    # TODO: Implement AI analysis
    return {
        "task_id": f"ai_{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "status": "accepted",
    }


@api_router.post("/ai/analyze/batch")
async def analyze_batch(project: str, rows: List[int], parallel: int = 3):
    """Batch AI analysis."""
    logger.info(
        "batch_analysis_requested",
        project=project,
        rows_count=len(rows),
    )

    # TODO: Implement batch analysis
    return {
        "task_id": f"batch_{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "status": "accepted",
        "total_rows": len(rows),
    }


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
