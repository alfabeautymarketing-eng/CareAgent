from typing import Optional, List, Dict, Any
from fastapi import APIRouter, HTTPException, Depends
from src.utils.logger import logger
from src.utils.logger import logger
from src.api.endpoints import sheets_service, certification_service
from src.services.invoice_service import get_invoice_service
from src.api.models import (
    InvoiceFormatRequest, InvoiceCreateRequest, InvoiceResponse,
    CollectDocumentsRequest
)

router = APIRouter(tags=["03. Инвойс"])

@router.post("/format-order", response_model=InvoiceResponse, summary="Нормализация листа заказа")
async def format_order_sheet(request: InvoiceFormatRequest):
    """Преобразует текстовые числа в числовые значения (Цена, Кол-во, Сумма)."""
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

@router.post("/create-full", response_model=InvoiceResponse, summary="Собрать полный инвойс")
async def create_full_invoice(request: InvoiceCreateRequest):
    """Создает сводный лист 'Для инвойса', объединяя данные из Ордера, Сертификации и Этикеток."""
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

@router.post("/collect-docs", summary="Сбор документов для инвойса")
async def collect_documents(request: CollectDocumentsRequest):
    """Собирает и копирует все необходимые документы для формирования инвойса."""
    try:
        inv_service = get_invoice_service(sheets_service)
        result = await inv_service.collect_and_copy_documents(
            spreadsheet_id=request.spreadsheet_id,
            target_sheet=request.target_sheet
        )
        return {"status": "success", "data": result}
    except Exception as e:
        logger.error("collect_documents_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/structure-353pp", summary="Подготовить документы (353пп)")
async def structure_documents_353pp(request: CollectDocumentsRequest):
    """Структурирует документы для подачи заявки по постановлению 353пп."""
    try:
        result = await certification_service.structure_documents_353pp(
            request.spreadsheet_id
        )
        return {"status": "success", "data": result}
    except Exception as e:
        logger.error("structure_353pp_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))
