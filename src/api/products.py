"""
Эндпоинты API для работы с товарами.
"""

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field
from typing import Optional, Any, Dict

from src.services.product_matcher import ProductMatcher
from src.utils.logger import logger

router = APIRouter()

class ProductMatchRequest(BaseModel):
    product_name: str = Field(..., description="Название нового товара для сопоставления")
    spreadsheet_id: str = Field(..., description="ID Google Таблицы, содержащей базовые товары")
    sheet_name: str = Field("Сертификация", description="Имя листа для поиска")

    class Config:
        title = "Запрос сопоставления товара"

@router.post("/match", summary="Сопоставить товар с базовыми товарами")
async def match_product(request: ProductMatchRequest) -> Dict[str, Any]:
    """
    Находит наиболее подходящий базовый товар для заданного имени товара
    с использованием Gemini AI и данных из Google Sheets.
    """
    logger.info(f"Received match request for '{request.product_name}'")
    
    matcher = ProductMatcher(request.spreadsheet_id, request.sheet_name)
    
    # 1. Fetch candidates
    candidates = matcher.fetch_base_products()
    if not candidates:
        logger.warning("No base products found or error fetching sheet.")
        # We return a valid response but with match_found=False, 
        # or we could error out. Returning no match seems safer for automation.
        return {
            "match_found": False, 
            "reasoning": "No base products found in the specified sheet."
        }
    
    # 2. Match
    result = matcher.find_best_match(request.product_name, candidates)
    
    if not result:
        raise HTTPException(status_code=500, detail="Failed to perform matching via AI service.")
        
    return result
