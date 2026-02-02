from typing import Optional, List, Dict, Any
from fastapi import APIRouter, HTTPException, Depends
from src.utils.logger import logger
from src.utils.logger import logger
from src.api.endpoints import sheets_service
from src.services.formula_service import get_formula_service
from src.api.models import (
    FormulaPriceDynamicsRequest, FormulaPriceCalcRequest, 
    FormulaAddYearRequest, FormulaResponse
)

router = APIRouter(tags=["04. Формулы"])

@router.post("/price-dynamics", response_model=FormulaResponse, summary="Пересчитать формулы динамики")
async def recalculate_price_dynamics_formulas(request: FormulaPriceDynamicsRequest):
    """Пересчитывает формулы в листе 'Динамика цены' (EXW, Закупка, DDP, Приросты)."""
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

@router.post("/price-calculation", response_model=FormulaResponse, summary="Обновить подгрузку цен")
async def update_price_calculation_formulas(request: FormulaPriceCalcRequest):
    """Обновляет формулы INDEX/MATCH в листе 'Расчет цены' из 'Динамики цены'."""
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

@router.post("/add-year-columns", response_model=FormulaResponse, summary="Добавить колонки нового года")
async def add_new_year_columns(request: FormulaAddYearRequest):
    """Добавляет 7 новых колонок для следующего года в лист 'Динамика цены'."""
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
