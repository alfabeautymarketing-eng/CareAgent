from typing import Optional, List, Dict
from fastapi import APIRouter, HTTPException, Depends
from src.utils.logger import logger
from src.api.endpoints import logging_service
from src.api.models import (
    LogArchiveRequest, LogResetRequest, LogRotationRequest, 
    LogEntryRequest, LogArchiveResponse, LogStatusResponse
)

router = APIRouter(prefix="/logs", tags=["Logs Management"])

@router.post("/archive", response_model=LogArchiveResponse, summary="Ежедневная архивация")
async def archive_logs_daily(request: LogArchiveRequest):
    """
    Архивирует все листы логов в ежемесячную таблицу в указанной папке Диска.
    """
    logger.info(
        "log_archive_requested",
        spreadsheet_id=request.spreadsheet_id,
        project_code=request.project_code,
        dry_run=request.dry_run
    )

    try:
        result = await logging_service.archive_logs_daily(
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

@router.post("/reset", response_model=LogArchiveResponse, summary="Сбросить листы логов")
async def reset_log_sheet(request: LogResetRequest):
    """
    Очищает данные в листе логов после архивации, сохраняя заголовки.
    """
    logger.info(
        "log_reset_requested",
        spreadsheet_id=request.spreadsheet_id,
        sheet_name=request.sheet_name,
        dry_run=request.dry_run
    )

    try:
        result = await logging_service.reset_daily_log_sheet(
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

@router.post("/rotation", response_model=LogArchiveResponse, summary="Ночная ротация логов")
async def midnight_log_rotation(request: LogRotationRequest):
    """
    Выполняет полный цикл ночной ротации: архивация + сброс листов.
    """
    logger.info(
        "log_rotation_requested",
        spreadsheet_id=request.spreadsheet_id,
        project_code=request.project_code,
        force=request.force,
        dry_run=request.dry_run
    )

    try:
        result = await logging_service.midnight_log_rotation(
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

@router.get("/status", response_model=LogStatusResponse, summary="Проверить статус архивов")
async def get_log_status(spreadsheet_id: str, project_code: str = "project"):
    """
    Возвращает информацию о последней архивации и количестве строк, ожидающих переноса.
    """
    logger.info(
        "log_status_requested",
        spreadsheet_id=spreadsheet_id,
        project_code=project_code
    )

    try:
        result = await logging_service.get_archive_status(
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

@router.post("/write", summary="Записать в лог")
async def write_log_entry(request: LogEntryRequest):
    """
    Записывает одиночную строку лога непосредственно в Google Таблицу.
    """
    try:
        result = await logging_service.write_log_entry(
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
