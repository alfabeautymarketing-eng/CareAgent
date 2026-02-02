from typing import Optional, List, Dict, Any
from pydantic import BaseModel, Field

class SortRequest(BaseModel):
    """Запрос на сортировку листа."""
    project: str = Field("Common", description="Код проекта")
    spreadsheet_id: str = Field(..., description="ID таблицы")
    sheet_name: str = Field(..., description="Имя листа")
    column_name: str = Field(..., description="Заголовок колонки для сортировки")
    ascending: bool = Field(True, description="По возрастанию?")

    class Config:
        title = "Запрос сортировки"

class StructureSortRequest(BaseModel):
    """Запрос на высокопроизводительную структурную сортировку."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    mode: str = Field(..., description="Режим: 'byManufacturer' (по производителю) или 'byPrice' (по цене)")

    class Config:
        title = "Запрос структурной сортировки"

class LoadFunctionsRequest(BaseModel):
    """Запрос на выполнение функций загрузки."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    project: str = Field("Common", description="Код проекта")

    class Config:
        title = "Запрос загрузки функций"

class OrderFilterRequest(BaseModel):
    """Запрос на фильтрацию этапов заказа."""
    spreadsheet_id: str = Field(..., description="ID таблицы")
    stage: str = Field("all", description="Этап: all, order, promotions, set, price")
    sheet_name: str = Field("Заказ", description="Имя листа")
    dry_run: bool = Field(False, description="Тестовый запуск")

    class Config:
        title = "Запрос фильтрации заказа"

class OrderFilterResponse(BaseModel):
    """Ответ фильтрации этапов заказа."""
    status: str = Field(..., description="Статус")
    stage: str = Field(..., description="Выбранный этап")
    visible_rows: int = Field(0, description="Видимых строк")
    hidden_rows: int = Field(0, description="Скрытых строк")
    hidden_columns: int = Field(0, description="Скрытых колонок")
    message: str = Field("", description="Сообщение")

    class Config:
        title = "Ответ фильтрации заказа"

class ReorderSheetsRequest(BaseModel):
    """Запрос на упорядочивание листов."""
    spreadsheet_id: str = Field(..., description="ID таблицы")

    class Config:
        title = "Запрос упорядочивания листов"
