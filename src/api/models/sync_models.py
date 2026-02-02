from typing import List, Optional
from pydantic import BaseModel, Field

class SyncRowRequest(BaseModel):
    """Запрос на синхронизацию одной строки."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field(..., description="Имя листа")
    row_number: int = Field(..., description="Номер строки")
    data: Optional[dict] = Field(None, description="Данные строки (опционально)")

    class Config:
        title = "Запрос синхронизации строки"

class SyncRangeRequest(BaseModel):
    """Запрос на синхронизацию диапазона."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field(..., description="Имя листа")
    start_row: int = Field(..., description="Начальная строка")
    end_row: int = Field(..., description="Конечная строка")

    class Config:
        title = "Запрос синхронизации диапазона"

class AddArticleRequest(BaseModel):
    """Запрос на добавление артикула."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field(..., description="Имя листа")
    article: str = Field(..., description="Артикул товара")

    class Config:
        title = "Запрос добавления артикула"

class DeleteArticlesRequest(BaseModel):
    """Запрос на удаление артикулов."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field(..., description="Имя листа")
    articles: List[str] = Field(..., description="Список артикулов")

    class Config:
        title = "Запрос удаления артикулов"

class SyncEventRequest(BaseModel):
    """Запрос события синхронизации."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    source_sheet: str = Field(..., description="Лист-источник")
    target_sheet: str = Field(..., description="Целевой лист")
    action: str = Field(..., description="Действие (update/delete)")
    row_key: str = Field(..., description="Ключ строки (Артикул)")
    details: Optional[str] = Field(None, description="Детали события")

    class Config:
        title = "Событие синхронизации"

class LogInitRequest(BaseModel):
    """Запрос инициализации лог-сессии."""
    spreadsheet_id: str = Field(..., description="ID таблицы")

    class Config:
        title = "Запрос инициализации логов"

class SyncBatchEventRequest(BaseModel):
    """Запрос пакета событий синхронизации."""
    events: List[SyncEventRequest]

    class Config:
        title = "Пакет событий синхронизации"

class SyncLogQueryParams(BaseModel):
    """Параметры фильтрации журнала синхронизации."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    limit: int = Field(200, description="Лимит записей")
    offset: int = Field(0, description="Смещение")
    order: str = Field("desc", description="Сортировка (asc/desc)")
    project: Optional[str] = Field(None, description="Фильтр по проекту")
    category: Optional[str] = Field(None, description="Фильтр по категории")
    status: Optional[str] = Field(None, description="Фильтр по статусу")
    row_key: Optional[str] = Field(None, description="Фильтр по артикулу")
    source: Optional[str] = Field(None, description="Фильтр по источнику")
    target: Optional[str] = Field(None, description="Фильтр по целевому листу")
    rule_id: Optional[str] = Field(None, description="Фильтр по ID правила")

    class Config:
        title = "Параметры журнала синхронизации"
