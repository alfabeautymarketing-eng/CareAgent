from typing import Optional, List, Dict, Any
from fastapi import APIRouter, HTTPException, Depends
from src.utils.logger import logger
from src.utils.logger import logger
from src.api.endpoints import sheets_service
from src.services.export_service import get_export_service, ExportType
from src.api.models import ExportRequest, ExportResponse

router = APIRouter(tags=["05. Экспорт"])

@router.post("/promotions", response_model=ExportResponse, summary="Экспорт акций")
async def export_promotions(request: ExportRequest):
    """Экспортирует данные по акциям из листа 'Заказ' в целевой документ проекта."""
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

@router.post("/sets", response_model=ExportResponse, summary="Экспорт наборов")
async def export_sets(request: ExportRequest):
    """Экспортирует данные о наборах товаров в целевой документ проекта."""
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
