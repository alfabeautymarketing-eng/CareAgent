from typing import List, Optional
from pydantic import BaseModel, Field

class AIAnalyzeRequest(BaseModel):
    """Запрос на анализ через AI."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field(..., description="Имя листа")
    row_number: Optional[int] = Field(None, description="Номер строки")

    class Config:
        title = "Запрос AI анализа"

class AIAnalyzeBatchRequest(BaseModel):
    """Запрос на пакетный анализ."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field("Главная", description="Имя листа")
    delay_between: float = Field(1.0, description="Задержка между строками")

    class Config:
        title = "Запрос пакетного AI анализа"

class AICheckServiceRequest(BaseModel):
    """Запрос проверки статуса AI."""
    project: str = Field("Common", description="Код проекта")

    class Config:
        title = "Запрос проверки AI"

class AIPdfAnalyzeRequest(BaseModel):
    """Запрос на анализ PDF."""
    pdf_url: str = Field(..., description="Ссылка на PDF")
    purpose: Optional[str] = Field(None, description="Цель анализа")
    application: Optional[str] = Field(None, description="Способ применения")

    class Config:
        title = "Запрос анализа PDF"

class AISimpleAnalyzeRequest(BaseModel):
    """Запрос на простой текстовый анализ."""
    product_name: str = Field(..., description="Наименование товара")
    purpose: Optional[str] = Field(None, description="Назначение")
    application: Optional[str] = Field(None, description="Применение")
    inci_text: Optional[str] = Field(None, description="INCI состав")

    class Config:
        title = "Запрос простого анализа"

class AIConfigureRequest(BaseModel):
    """Запрос на конфигурацию настроек AI."""
    api_key: Optional[str] = Field(None, description="API ключ Gemini")
    model: Optional[str] = Field(None, description="Имя модели (например, gemini-1.5-pro)")
    spreadsheet_id: Optional[str] = Field(None, description="ID таблицы")
    model_name: Optional[str] = Field(None, description="Альтернативное имя модели")
    system_instruction: Optional[str] = Field(None, description="Системная инструкция")

    class Config:
        title = "Запрос настройки ИИ"
