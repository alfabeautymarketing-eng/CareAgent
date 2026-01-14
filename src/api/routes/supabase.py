from typing import Any

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from src.services.supabase_service import get_supabase_service

router = APIRouter()

class SqlQueryRequest(BaseModel):
    query: str

class TableDataRequest(BaseModel):
    data: list[dict[str, Any]]

@router.get("/test", summary="Test Supabase connection")
async def test_connection(spreadsheet_id: str | None = None):
    service = get_supabase_service()
    if service.is_configured():
        # Try a lightweight call
        try:
            tables = service.get_tables()
            return {"status": "ok", "message": "Connected", "tables": len(tables)}
        except Exception as e:
            return {"status": "error", "message": str(e)}
    else:
        raise HTTPException(status_code=503, detail="Supabase not configured on server")

@router.get("/tables", summary="Get list of tables")
async def get_tables():
    service = get_supabase_service()
    if not service.is_configured():
        raise HTTPException(status_code=503, detail="Supabase not configured")
    return {"tables": service.get_tables()}

@router.post("/query", summary="Execute SQL query")
async def execute_query(request: SqlQueryRequest):
    service = get_supabase_service()
    if not service.is_configured():
        raise HTTPException(status_code=503, detail="Supabase not configured")

    try:
        result = service.execute_query(request.query)
        return result
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/table/{table_name}", summary="Get data from table")
async def get_table_data(table_name: str, limit: int = 100):
    service = get_supabase_service()
    if not service.is_configured():
        raise HTTPException(status_code=503, detail="Supabase not configured")

    try:
        data = service.get_table_data(table_name, limit)
        return {"data": data}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.post("/table/{table_name}", summary="Insert data into table")
async def insert_table_data(table_name: str, request: TableDataRequest):
    service = get_supabase_service()
    if not service.is_configured():
        raise HTTPException(status_code=503, detail="Supabase not configured")

    # Not implemented fully in service yet, but let's add a placeholder
    # service.insert_data(table_name, request.data)
    return {"status": "not_implemented_yet", "message": "Write operations pending implementation"}
