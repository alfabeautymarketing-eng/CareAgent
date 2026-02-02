import json
from datetime import datetime
from typing import Optional, List
from fastapi import APIRouter, HTTPException, BackgroundTasks, Depends
from src.utils.logger import logger
from src.api.endpoints import ai_service, logging_service, sheets_service
from src.services.gemini_client import set_runtime_config, get_gemini_client
from src.api.models import (
    AIAnalyzeRequest, AIAnalyzeBatchRequest, AICheckServiceRequest,
    AIPdfAnalyzeRequest, AISimpleAnalyzeRequest, AIConfigureRequest
)

router = APIRouter(prefix="/ai", tags=["AI Services"])

@router.get("/status", summary="Статус AI сервиса")
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

@router.get("/settings", summary="Настройки AI")
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

@router.post("/configure", summary="Настроить AI")
async def configure_ai(request: AIConfigureRequest):
    """
    Настраивает параметры Gemini AI (ключ, модель) "на лету" без перезагрузки сервера.
    """
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

@router.get("/categories", summary="Категории ТН ВЭД")
async def get_ai_categories():
    """Возвращает список доступных категорий продукции для классификации."""
    return {
        "categories": ai_service.get_available_categories()
    }

@router.post("/analyze/pdf", summary="Анализ PDF документа")
async def analyze_pdf(request: AIPdfAnalyzeRequest):
    """
    Анализирует PDF документ (состав, описание) через AI без привязки к таблице.
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

@router.post("/analyze/simple", summary="Тест AI (быстрый анализ)")
async def analyze_simple(request: AISimpleAnalyzeRequest):
    """
    Простая классификация продукта по названию и назначению для проверки работы API.
    """
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
        
        # Handle case when Gemini returns a list instead of dict
        if isinstance(result, list):
            if len(result) > 0 and isinstance(result[0], dict):
                result = result[0]
            else:
                raise ValueError(f"Unexpected list response from Gemini: {result}")
        
        logger.info("simple_analysis_success",
                   product_type=result.get("product_type"))

        return {
            "status": "success",
            "data": result
        }
    except Exception as e:
        logger.error("simple_analysis_failed", error=str(e))
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/analyze/row", summary="Анализ строки AI")
async def analyze_row(request: AIAnalyzeRequest):
    """
    Запускает AI анализ для конкретной строки таблицы (парсинг INCI, классификация).
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
            # Note: logging_service dependency needs to be handled properly
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

@router.post("/analyze/batch", summary="Пакетный анализ AI")
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
