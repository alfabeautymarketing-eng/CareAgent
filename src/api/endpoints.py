"""
API endpoints for sync, price processing, AI analysis.
"""

from datetime import datetime
from typing import Optional, List, Any

from fastapi import APIRouter, HTTPException, BackgroundTasks
from pydantic import BaseModel

from src.utils.logger import logger

api_router = APIRouter()

from src.services.sheets import SheetsService
from src.services.sync import SyncService
from src.services.sorting import SortingService
from src.services.logging import LoggingService
from src.services.ai import AIService, get_ai_service

sheets_service = SheetsService()
logging_service = LoggingService(sheets_service)
sync_service = SyncService(logging_service) # Pass logger to sync service
sorting_service = SortingService()
ai_service = get_ai_service()

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
    row_key: Optional[str] = None

class LogInitRequest(BaseModel):
    """Request to initialize logs."""
    spreadsheet_id: str

class SyncBatchEventRequest(BaseModel):
    """Request for batch onEdit events."""
    spreadsheet_id: str
    events: List[SyncEventRequest]


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
        result = sync_service.sync_row(spreadsheet_id, request.project, request.article, request.source_sheet)
        return {
            "task_id": f"row_{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "status": "success",
            "message": f"Row sync processed for {request.article}",
            "details": result
        }
    except Exception as e:
        logger.error("sync_row_endpoint_failed", error=str(e))
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
            "row_key": request.row_key
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
                "row_key": event.row_key
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
        result = sorting_service.sort_sheets(request.spreadsheet_id, request.mode)
        return {
            "status": "success",
            "message": f"Sorted {len(result['sheets_processed'])} sheets",
            "details": result
        }
    except Exception as e:
        logger.error("structure_sort_failed", error=str(e))
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

# Project spreadsheet IDs mapping
PROJECT_IDS = {
    "199Np7xsBiBRQih5_tlUdpt6EmkfRGjZAhTvKm4Ua0Q6XEaMtvAmQUn0g": "MT",
    "1sTgZa-n1aP7oIhyQfPeN8QDgDNnCubqMWAd-TKjKpJXWsQm_ZhXnojPD": "SS",
    "1DJvK1vUT2OTubN0TLdZvsgYMSYByLHl8xTsus3K-KJ-VtJxgGnSw5Ih8": "SK",
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
    "MAIN": {"fn_by_project": {"SK": "processSkPriceSheet", "SS": "processSsPriceSheet", "MT": "processMtMainPrice"}},
    "TESTER": {"fn_by_project": {"MT": "processMtTesterPrice"}},
    "SAMPLES": {"fn_by_project": {"MT": "processMtSamplesPrice"}},
    "PROBES": {"fn_by_project": {"SK": "processSkPriceProbes"}},
    "STOCKS": {"fn_by_project": {"SK": "loadSkStockData", "SS": "loadSsStockData", "MT": "loadMtStockData"}},
    "NEW_PRICE_YEAR": {"fn": "addNewYearColumnsToPriceDynamics"},
    "SORT_MANUFACTURER": {"fn": "sortByManufacturer"},
    "SORT_PRICE": {"fn": "sortByPrice"},
    "STAGE_ALL": {"fn": "showAllOrderData"},
    "STAGE_ORDER": {"fn": "showOrderStage"},
    "STAGE_PROMOTIONS": {"fn": "showPromotionsStage"},
    "STAGE_SET": {"fn": "showSetStage"},
    "STAGE_PRICE": {"fn": "showPriceStage"},
}

# Static menu groups from legacy GAS config (without project-specific items)
BASE_MENU_GROUPS: List[dict] = [
    {
        "title": "📦 Выгрузка",
        "items": [
            {"label": "Выгрузить Акции", "function_name": "exportPromotions"},
            {"label": "Выгрузить Наборы", "function_name": "exportSets"},
        ],
    },
    {
        "title": "🚚 Поставка",
        "items": [
            {"label": "Форматировать лист 'Ордер'", "function_name": "formatOrderSheet"},
            {"separator": True},
            {"label": "1. Создать лист 'Для инвойса'", "function_name": "createFullInvoice"},
            {"label": "2. Собрать документы", "function_name": "collectAndCopyDocuments"},
        ],
    },
    {
        "title": "🔬 Сертификация",
        "items": [
            {"label": "Лист новинки", "function_name": "createNewsSheetFromCertification"},
            {"separator": True},
            {"label": "Создать заявку протоколы (353пп)", "function_name": "generateProtocols_353pp"},
            {"label": "Создать заявку ДС (353пп)", "function_name": "generateDsLayouts_353pp"},
            {"label": "Собрать документы для заявки (353пп)", "function_name": "structureDocuments_353pp"},
            {"separator": True},
            {"label": "Посчитать спирты", "function_name": "calculateAndAssignSpiritNumbers"},
            {"label": "Создать Макеты спирты", "function_name": "generateSpiritProtocols"},
            {"separator": True},
            {"label": "Пересчитать каскады (Сертификация)", "function_name": "runManualCascadeOnCertification"},
        ],
    },
    {
        "title": "⚙️ Синхронизация",
        "items": [
            {
                "submenu": "⚙️ Настройки",
                "items": [
                    {"label": "🔄 Обновить триггеры", "function_name": "setupTriggers"},
                    {"label": "📝 Настроить правила", "function_name": "showSyncConfigDialog"},
                    {"label": "📄 Внешние документы", "function_name": "showExternalDocManagerDialog"},
                    {"label": "📋 Создать/Пересоздать журнал синхро", "function_name": "recreateLogSheet"},
                    {"label": "🧹 Очистить журнал", "function_name": "quickCleanLogSheet"},
                ],
            },
            {"separator": True},
            {"label": "➕ Добавить артикул", "function_name": "addArticleManually"},
            {"label": "❌ Удалить артикул", "function_name": "deleteSelectedRowsWithSync"},
            {"separator": True},
            {"label": "🔄 Синхронизировать строку", "function_name": "syncSelectedRow"},
            {"label": "🔄 Синхронизировать всю таблицу", "function_name": "runFullSync"},
        ],
    },
    {
        "title": "🤖 Агент",
        "items": [
            {"label": "🔑 Ввести API ключ", "function_name": "setupGeminiComplete"},
            {"label": "📋 Показать настройки", "function_name": "showGeminiSettings"},
            {"label": "🟢 Проверить сервис", "function_name": "menuCheckService"},
            {"separator": True},
            {"label": "Анализировать выбранную строку", "function_name": "menuAnalyzeSelected"},
            {"label": "Анализировать пустые строки", "function_name": "menuAnalyzeEmpty"},
            {"separator": True},
            {"label": "Показать категории", "function_name": "menuShowCategories"},
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

    primary_group = _build_primary_menu(project)
    if primary_group:
        registry.append(primary_group)

    stages_group = _build_order_stages_menu(project)
    if stages_group:
        registry.append(stages_group)

    for group in BASE_MENU_GROUPS:
        registry.append(_clone_base_group(group))

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
    "Правила синхро",
    "Журнал синхро",
    "Журнал логов",
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
