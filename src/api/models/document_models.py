from typing import Optional
from pydantic import BaseModel, Field

class CollectDocumentsRequest(BaseModel):
    """Запрос на сбор документов."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    target_sheet: str = Field("Для инвойса", description="Имя целевого листа")

    class Config:
        title = "Запрос сбора документов"

class CertificationNewsRequest(BaseModel):
    """Запрос на создание листа новых товаров."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    source_sheet: str = Field("Сертификация", description="Имя исходного листа")
    target_sheet: str = Field("New sert", description="Имя нового листа")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос листа новинок"

class CertificationSpiritsRequest(BaseModel):
    """Запрос на расчет спиртовых номеров."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field("Сертификация", description="Имя листа")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос расчета спиртов"

class CertificationProtocolsRequest(BaseModel):
    """Запрос на генерацию протоколов/макетов."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    protocol_type: str = Field("353pp", description="Тип (например, 353пп)")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос генерации протоколов"

class CertificationResponse(BaseModel):
    """Ответ по операциям сертификации."""
    status: str = Field(..., description="Статус")
    rows_affected: int = Field(0, description="Количество затронутых строк")
    sheet_name: str = Field("", description="Имя листа")
    message: str = Field("", description="Сообщение")

    class Config:
        title = "Ответ сертификации"
