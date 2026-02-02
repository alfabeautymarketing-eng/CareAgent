from typing import Optional, List, Dict, Any
from fastapi import APIRouter, HTTPException, Depends
from src.utils.logger import logger
from src.utils.logger import logger
from src.api.endpoints import sheets_service, certification_service
from src.services.cascade_processor import get_cascade_processor
from src.api.models import (
    CascadeProcessRequest, CascadeProcessResponse,
    CertificationNewsRequest, CertificationSpiritsRequest, 
    CertificationProtocolsRequest, CertificationResponse
)

router = APIRouter(tags=["02. Сертификация"])

@router.post("/cascade/process", response_model=CascadeProcessResponse, summary="Запустить каскадную обработку")
async def process_cascade(request: CascadeProcessRequest):
    """
    Обрабатывает каскадные правила для листа сертификации.
    """
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

@router.post("/cascade/recalculate-all", summary="Пересчитать все каскады")
async def recalculate_all_cascades(request: CascadeProcessRequest):
    """Recalculate cascades for all rows in certification sheet."""
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
            "dry_run": result.get("dry_run", False)
        }
    except Exception as e:
        logger.error("cascade_recalculate_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/news-sheet", response_model=CertificationResponse, summary="Создать лист New Sert")
async def create_news_sheet(request: CertificationNewsRequest):
    """Создает лист 'New sert' на основе данных из 'Сертификация'."""
    try:
        result = await certification_service.create_news_sheet(
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

@router.post("/spirits/calculate", response_model=CertificationResponse, summary="Расчет спиртов")
async def calculate_spirits(request: CertificationSpiritsRequest):
    """Рассчитывает и назначает спиртовые номера строкам сертификации."""
    try:
        result = await certification_service.calculate_spirit_numbers(
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

@router.post("/protocols-353pp", response_model=CertificationResponse, summary="Генерация протоколов 353пп")
async def generate_protocols_353pp(request: CertificationProtocolsRequest):
    """Генерирует протоколы сертификации по постановлению 353пп."""
    try:
        result = await certification_service.generate_protocols(
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

@router.post("/ds-layouts", response_model=CertificationResponse, summary="Генерация макетов ДС")
async def generate_ds_layouts(request: CertificationProtocolsRequest):
    """Генерирует макеты деклараций о соответствии (ДС)."""
    try:
        result = await certification_service.generate_ds_layouts(
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
