from typing import Optional
from pydantic import BaseModel, Field

class PriceProcessRequest(BaseModel):
    """Запрос на обработку файла с прайсами."""
    spreadsheet_id: str = Field(..., description="ID таблицы, в которой нужно обновить цены")
    mode: str = Field("main", description="Режим: main (основной), tester, samples, probes")
    source_doc_id: Optional[str] = Field(None, description="ID исходного документа с прайсом")
    dry_run: bool = Field(False, description="Тестовый запуск без изменения данных")

    class Config:
        title = "Запрос на обработку прайса"

class StockLoadRequest(BaseModel):
    """Запрос на загрузку остатков."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    source_doc_id: Optional[str] = Field(None, description="ID исходного документа с остатками")

    class Config:
        title = "Запрос на загрузку остатков"

class StockLoadResponse(BaseModel):
    """Ответ после загрузки остатков."""
    status: str = Field(..., description="Статус выполнения")
    message: str = Field(..., description="Сообщение")
    updated_rows: int = Field(0, description="Количество обновленных строк")
    source_doc: str = Field("", description="ID использованного документа-источника")

    class Config:
        title = "Ответ загрузки остатков"
