from typing import Optional, List, Dict
from fastapi import APIRouter, HTTPException, BackgroundTasks, Depends
from src.utils.logger import logger
from src.api.models import ExternalDocAddRequest
from src.api.endpoints import (
    sheets_service, meta_cache_service, external_docs_service, PROJECT_IDS
)
from src.config.project_menus import get_project_menu

# Constants needed for meta
from src.api.endpoints import PROJECT_IDS

router = APIRouter(prefix="/meta", tags=["System Meta"])
docs_router = APIRouter(prefix="/external-docs", tags=["External Documents"])

@router.get("/menu/config", summary="Получить меню по spreadsheet_id (для GAS)")
async def get_menu_config_by_spreadsheet(spreadsheet_id: str):
    """
    Get menu configuration by spreadsheet ID (for Google Apps Script).
    """
    try:
        project = PROJECT_IDS.get(spreadsheet_id)
        if not project:
            logger.warning("Unknown spreadsheet_id", spreadsheet_id=spreadsheet_id)
            raise ValueError(f"Unknown spreadsheet_id: {spreadsheet_id}")

        config = get_project_menu(project)

        menus = []
        for group_cfg in config.get("menu_groups", []):
            group_items = []
            for item_cfg in group_cfg.get("items", []):
                if item_cfg.get("separator"):
                    group_items.append({"separator": True})
                group_items.append({
                    "label": item_cfg.get("label"),
                    "function_name": item_cfg.get("function_name"),
                })
            
            menus.append({
                "title": group_cfg.get("title"),
                "items": group_items
            })

        return {
            "project": project,
            "menus": menus
        }
    except Exception as e:
        logger.error("menu_config_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/{spreadsheet_id}/sheets", summary="Список листов таблицы")
async def get_meta_sheets(spreadsheet_id: str):
    """Возвращает названия всех листов в указанной Google Таблице."""
    try:
        sh = sheets_service.gc.open_by_key(spreadsheet_id)
        return {"sheets": [s.title for s in sh.worksheets()]}
    except Exception as e:
        logger.error(f"Failed to get sheets for {spreadsheet_id}: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/{spreadsheet_id}/{sheet_name}/headers", summary="Список заголовков листа")
async def get_meta_headers(spreadsheet_id: str, sheet_name: str):
    """Возвращает список всех названий колонок (заголовков) на указанном листе."""
    try:
        headers = sheets_service.get_worksheet_headers(spreadsheet_id, sheet_name)
        return {"headers": headers}
    except Exception as e:
        logger.error(f"Failed to get headers for {spreadsheet_id}/{sheet_name}: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/discovery", summary="Карта заголовков системы")
async def get_meta_discovery():
    """Строит и возвращает карту общих заголовков во всей экосистеме проектов."""
    return meta_cache_service.get_discovery_map()

@router.post("/index", summary="Запустить переиндексацию метаданных")
async def trigger_meta_index(background_tasks: BackgroundTasks, spreadsheet_id: Optional[str] = None):
    """
    Запускает процесс обновления кэша метаданных. 
    Если указан spreadsheet_id, обновится только эта таблица.
    Иначе — все проекты из PROJECT_IDS.
    """
    if spreadsheet_id:
        targets = [spreadsheet_id]
    else:
        # Get all unique IDs from PROJECT_IDS
        targets = list(PROJECT_IDS.keys())
    
    background_tasks.add_task(meta_cache_service.build_global_index, targets)
    return {"status": "queued", "targets": len(targets)}

@router.get("/index/status", summary="Статус кэша метаданных")
async def get_meta_index_status():
    """Возвращает текущее состояние и статистику кэша метаданных."""
    return {
        "updated_at": meta_cache_service.index.get("updated_at"),
        "total_spreadsheets": len(meta_cache_service.index.get("spreadsheets", {})),
        "total_headers": len(meta_cache_service.index.get("headers", {})),
        "spreadsheets": meta_cache_service.get_indexed_spreadsheets()
    }

@router.get("/search", summary="Поиск по заголовкам")
async def search_meta_header(q: str):
    """Ищет листы, содержащие указанный заголовок колонки."""
    return meta_cache_service.search_header(q)

# External Docs Endpoints
@docs_router.get("", summary="Список внешних документов")
async def list_external_docs():
    """Возвращает список всех зарегистрированных внешних Google Таблиц."""
    return {"docs": external_docs_service.list_docs()}

@docs_router.post("", summary="Добавить внешний документ")
async def add_external_doc(request: ExternalDocAddRequest):
    """Регистрирует новую внешнюю Google Таблицу в системе."""
    try:
        doc = external_docs_service.add_doc(request.name, request.doc_id)
        return {"status": "ok", "doc": doc}
    except Exception as e:
        logger.error(f"Failed to add external doc: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@docs_router.delete("/{doc_id}", summary="Удалить внешний документ")
async def remove_external_doc(doc_id: str):
    """Удаляет регистрацию внешней Google Таблицы."""
    if external_docs_service.remove_doc(doc_id):
        return {"status": "ok"}
    raise HTTPException(status_code=404, detail="Document not found")
