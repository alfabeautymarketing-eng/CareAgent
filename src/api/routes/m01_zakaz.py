from datetime import datetime
from typing import Optional, List, Dict, Any
from fastapi import APIRouter, HTTPException, BackgroundTasks, Depends
from src.utils.logger import logger
from src.api.endpoints import (
    sheets_service, sorting_service, price_processor, 
    stock_processor, logging_service, PROJECT_IDS, MENU_CONFIGS
)
from src.api.models import (
    SortRequest, StructureSortRequest, LoadFunctionsRequest,
    PriceProcessRequest, StockLoadRequest, StockLoadResponse,
    OrderFilterRequest, OrderFilterResponse, TaskStatusResponse
)


router = APIRouter(tags=["01. Заказ"])

@router.post("/sort", summary="Сортировка по колонке")
async def sort_sheet(request: SortRequest):
    """Выполняет сортировку указанного листа по заданному заголовку колонки."""
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

@router.post("/sort/structure", summary="Структурная сортировка листов")
async def sort_structure(request: StructureSortRequest):
    """
    Высокопроизводительная группировка и сортировка нескольких листов (Заказ, Динамика цены, Расчет цены).
    """
    logger.info("structure_sort_requested", spreadsheet_id=request.spreadsheet_id, mode=request.mode)
    
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
        project = PROJECT_IDS.get(request.spreadsheet_id, "MT")
        config = MENU_CONFIGS.get(project, {})
        group_header_color = config.get("group_header_color")
        
        result = sorting_service.sort_sheets(
            request.spreadsheet_id, 
            request.mode,
            group_header_color=group_header_color
        )
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

@router.post("/load-functions", summary="Выполнить логику при загрузке")
async def run_load_functions(request: LoadFunctionsRequest):
    """Запускает тяжелую логику обновления формул и данных при открытии таблицы."""
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

@router.post("/price/process/{project}", summary="Обработать прайс поставщика")
async def process_price(project: str, request: PriceProcessRequest, background_tasks: BackgroundTasks):
    try:
        if request.dry_run:
            return await price_processor.process(
                project=project, mode=request.mode, spreadsheet_id=request.spreadsheet_id,
                source_doc_id=request.source_doc_id, dry_run=True
            )
        background_tasks.add_task(
            price_processor.process, project=project, mode=request.mode, 
            spreadsheet_id=request.spreadsheet_id, source_doc_id=request.source_doc_id, dry_run=False
        )
        return {"status": "queued", "message": f"Processing {project.upper()} {request.mode} started"}
    except Exception as e:
        logger.error("price_process_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/stocks/load", response_model=StockLoadResponse, summary="Загрузить остатки")
async def load_stocks(request: StockLoadRequest):
    """Загружает данные об остатках из листа '-остатки' в лист 'Заказ'."""
    project_code = PROJECT_IDS.get(request.spreadsheet_id, "MT").lower()
    try:
        result = await stock_processor.process(
            spreadsheet_id=request.spreadsheet_id,
            project=project_code,
            source_doc_id=request.source_doc_id
        )
        return StockLoadResponse(**result)
    except Exception as e:
        logger.error("stock_load_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/order/filter", response_model=OrderFilterResponse, summary="Фильтр этапов заказа")
async def order_stage_filter(request: OrderFilterRequest):
    """Переключает видимость колонок и строк в зависимости от выбранного этапа заказа."""
    logger.info("order_filter_requested", stage=request.stage, spreadsheet_id=request.spreadsheet_id)
    try:
        from src.services.order_service import get_order_service
        service = get_order_service(sheets_service)
        result = await order_service.apply_stage_filter(
            request.spreadsheet_id,
            request.stage,
            sheet_name=request.sheet_name,
            dry_run=request.dry_run
        )
        return OrderFilterResponse(**result)
    except Exception as e:
        logger.error("order_filter_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))
